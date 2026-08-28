#include <gtest/gtest.h>
#include <iostream>
#include "InnerLogger.hpp"


/** Controlling outer class. Maintains lifetime also for inner classes. */
class ControllingClass :
    public CComObjectRootEx<CComMultiThreadModel>, // also compatible with STA
    public CComCoClass<ControllingClass>, // no CLSID needed
    public IMyMessage {
public:
    ControllingClass() {
        std::wcout << L"ControllingClass ctor" << std::endl;
        CComAggObject<InnerLogger>::CreateInstance(this, &m_logger);
        assert(m_logger);
        m_logger->AddRef();
        s_obj_count++;
    }
    ~ControllingClass() {
        std::wcout << L"ControllingClass dtor" << std::endl;

        m_logger->Release();
        m_logger = nullptr;
        s_obj_count--;
    }

    HRESULT STDMETHODCALLTYPE ShowMessage(BSTR /*msg*/) override {
        return E_NOTIMPL;
    }

    BEGIN_COM_MAP(ControllingClass)
        COM_INTERFACE_ENTRY(IMyMessage)
#ifdef _WIN32
        COM_INTERFACE_ENTRY_AGGREGATE(__uuidof(IMyLogger), m_logger)
#else
        COM_INTERFACE_ENTRY_AGGREGATE(IMyLogger, m_logger)
#endif
    END_COM_MAP()

    static inline ULONG s_obj_count = 0;

private:
    CComAggObject<InnerLogger>* m_logger = nullptr; // inner class
};



TEST(AggregationTests, TestBasicAggregation) {
    {
        CComPtr<IMyLogger> inner;
        {
            CComObject<ControllingClass>* obj = nullptr;
            CHECK(CComObject<ControllingClass>::CreateInstance(&obj));
            CHECK(obj->QueryInterface(__uuidof(IMyLogger), (void**)&inner));
            ASSERT_NE(inner, nullptr);
        }
        ASSERT_EQ(GetRefCount(*inner), 1u);

        inner->Log(CComBSTR(L"Hi there"));

        CComPtr<IMyMessage> outer;
        CHECK(inner.QueryInterface(&outer));
        ASSERT_NE(outer, nullptr);
        ASSERT_EQ(GetRefCount(*outer), 2u);

        inner.Release();
        ASSERT_EQ(GetRefCount(*outer), 1u);

        CHECK(outer.QueryInterface(&inner));
        ASSERT_EQ(GetRefCount(*inner), 2u);

        inner->Log(CComBSTR(L"Hi again"));
    }

    ASSERT_EQ(InnerLogger::s_obj_count, 0u);
    ASSERT_EQ(ControllingClass::s_obj_count, 0u);
}
