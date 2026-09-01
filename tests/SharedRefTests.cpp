#include <gtest/gtest.h>
#include <iostream>
#include "InnerLogger.hpp"
#include <SharedRef.hpp>


TEST(SharedRefTests, CreateDestroy) {
    {
        CComPtr<IUnknown> obj(new SharedRef<InnerLogger>());
        ASSERT_EQ(GetRefCount(*obj), 1u);
    }
    auto count = SharedRef<InnerLogger>::ObjectCount();
    ASSERT_EQ(count, 0u);
    ASSERT_EQ(InnerLogger::s_obj_count, 0u);
}

TEST(SharedRefTests, StrongOutliveWeak) {
    {
        CComPtr<IUnknown> strong;
        CComPtr<IWeakRef> weak;
        {
            CComPtr<IUnknown> obj(new SharedRef<InnerLogger>());
            ASSERT_EQ(GetRefCount(*obj), 1u);
            CHECK(obj.QueryInterface(&strong));
            CHECK(obj.QueryInterface(&weak));
        }
        ASSERT_EQ(GetRefCount(*strong), 1u);
        ASSERT_EQ(GetRefCount(*weak), 2u);

        {
            // cast from IWeakRef to IPluginLogger
            CComPtr<IMyLogger> inner;
            CHECK(weak.QueryInterface(&inner));
            ASSERT_EQ(GetRefCount(*inner), 2u);
            ASSERT_EQ(GetRefCount(*strong), 2u);

            inner->Log(CComBSTR(L"Hi there"));
        }
        ASSERT_EQ(GetRefCount(*strong), 1u);

        {
            // cast from IWeakRef to IUnknown
            CComPtr<IUnknown> obj;
            CHECK(weak.QueryInterface(&obj));
            ASSERT_EQ(GetRefCount(*strong), 2u);

            // cast from IUnknown to IPluginLogger
            CComPtr<IMyLogger> inner;
            CHECK(obj.QueryInterface(&inner));
            ASSERT_EQ(GetRefCount(*inner), 3u);
            ASSERT_EQ(GetRefCount(*strong), 3u);

            // cast from IPluginLogger to IUnknown
            obj.Release();
            ASSERT_EQ(GetRefCount(*strong), 2u);
            CHECK(inner.QueryInterface(&obj));
            ASSERT_EQ(GetRefCount(*inner), 3u);
            ASSERT_EQ(GetRefCount(*strong), 3u);

            // cast from IPluginLogger to IWeakRef
            weak.Release();
            ASSERT_EQ(GetRefCount(*strong), 3u);
            CHECK(inner.QueryInterface(&weak));
            ASSERT_EQ(GetRefCount(*weak), 2u);
            ASSERT_EQ(GetRefCount(*strong), 3u);
        }
        ASSERT_EQ(GetRefCount(*strong), 1u);

        // let strong pointer outlive weak pointer
        weak.Release();
        ASSERT_EQ(GetRefCount(*strong), 1u);
    }
    auto count = SharedRef<InnerLogger>::ObjectCount();
    ASSERT_EQ(count, 0u);
    ASSERT_EQ(InnerLogger::s_obj_count, 0u);
}

TEST(SharedRefTests, WeakOutliveStrong) {
    {
        CComPtr<IUnknown> strong;
        CComPtr<IWeakRef> weak;
        {
            CComPtr<IUnknown> obj(new SharedRef<InnerLogger>());
            ASSERT_EQ(GetRefCount(*obj), 1u);
            CHECK(obj.QueryInterface(&strong));
            CHECK(obj.QueryInterface(&weak));
        }
        ASSERT_EQ(GetRefCount(*strong), 1u);
        ASSERT_EQ(GetRefCount(*weak), 2u);

        // let weak pointer outlive strong pointer
        strong.Release();
        ASSERT_EQ(GetRefCount(*weak), 1u);

        {
            // verify that Resolve fail when use-count=0
            CComPtr<IUnknown> obj;
            HRESULT hr = weak.QueryInterface(&obj);
            ASSERT_EQ(hr, E_NOT_SET);
        }
    }
    auto count = SharedRef<InnerLogger>::ObjectCount();
    ASSERT_EQ(count, 0u);
    ASSERT_EQ(InnerLogger::s_obj_count, 0u);
}

TEST(SharedRefTests, Test_SharedRef_Comparison)
{
    SharedRef<InnerLogger> *sharedRef = new SharedRef<InnerLogger>();
    InnerLogger* innerLogger = sharedRef->Internal();
    const IUnknownPtr unknownPtr(sharedRef);

    EXPECT_TRUE(unknownPtr == sharedRef);
    EXPECT_TRUE(sharedRef == unknownPtr);

    // test bool _com_ptr_t::operator==(T* p) const
    EXPECT_TRUE(unknownPtr == innerLogger);

    // test bool _com_ptr_t::operator==(T* p, const _com_ptr_t& _This)
    EXPECT_TRUE(innerLogger == unknownPtr);
}
