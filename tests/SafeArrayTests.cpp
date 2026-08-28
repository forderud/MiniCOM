#include <array>
#include <AppAPI/ComSupport.hpp>
#include <gtest/gtest.h>

#ifdef _WIN32
template <unsigned int N>
CComSafeArray<BYTE> CreateSafeArray (const CComSafeArrayBound (&bounds)[N]) {
    // assert that CComSafeArray is no larger than a SAFEARRAY pointer
    static_assert(sizeof(SAFEARRAY*) == sizeof(CComSafeArray<BYTE>), "CComSafeArray size mismatch");

    CComSafeArray<BYTE> arr(bounds, N);

    // verify that dimensions are reversed
    for (size_t i = 0; i < N; ++i)
        assert(arr.m_psa->rgsabound[i].cElements == bounds[N-1-i].cElements);

    // verify that array type is VT_UI1
    {
        VARTYPE type = 0;
        CHECK(SafeArrayGetVartype(arr, &type));
        assert(type == VT_UI1);
        type = arr.GetType();
        assert(type == VT_UI1);
        type = reinterpret_cast<VARTYPE*>(arr.m_psa)[-2]; // type at offset -0x04 (REF: http://source.winehq.org/git/wine.git/blob/HEAD:/dlls/oleaut32/safearray.c)
        assert(type == VT_UI1);
    }

    return arr;
}


/** Creates a SafeArray object that points to existing data.
    WARNING: Returned object will crash as destruction unless FADF_AUTO, FADF_STATIC or FADF_EMBEDDED is set. */
template <unsigned int N>
SAFEARRAY * CreateWeakSafeArray(std::array<BYTE, N> & buffer, USHORT extra_flags) {
    assert(extra_flags & (FADF_AUTO | FADF_STATIC | FADF_EMBEDDED));
    SAFEARRAY * sa_obj = nullptr;
    CHECK(SafeArrayAllocDescriptorEx(VT_UI1, 1, &sa_obj));
    sa_obj->cbElements = 1; // element size
    sa_obj->fFeatures |= extra_flags;
    sa_obj->rgsabound[0] = { static_cast<ULONG>(buffer.size()), 0 };
    {
        CHECK(SafeArrayLock(sa_obj));
        sa_obj->pvData = buffer.data();
        CHECK(SafeArrayUnlock(sa_obj));
    }
    return sa_obj;
}


TEST(SafeArrayTests, VerifyThat_FADF_STATIC_Clears_NeitherDeletesDataNorDescriptor) {
    std::array<BYTE, 8> buffer = { 0, 1, 2, 3, 4, 5, 6, 7 }; // buffer[i] = i

    SAFEARRAY * sa_ptr = nullptr;
    {
        // create an array
        CComSafeArray<BYTE> arr;
        arr.Attach(CreateWeakSafeArray(buffer, FADF_STATIC)); // tag data to be cleared but not deleted at destruction

        sa_ptr = arr.m_psa;
        // SafeArrayDestroy in "arr" dtor clears the data, but does not delete data or descriptor
    }

    // verify that array is cleared but NOT deleted
    ASSERT_EQ(sa_ptr->cDims, 1);
    ASSERT_EQ(sa_ptr->cbElements, 1u);
    for (size_t i = 0; i < buffer.size(); ++i)
        ASSERT_EQ(reinterpret_cast<BYTE*>(sa_ptr->pvData)[i], 0);

    // manually delete SAFEARRAY descriptor
    CHECK(SafeArrayDestroyDescriptor(sa_ptr));
}


TEST(SafeArrayTests, VerifyThat_FADF_AUTO_NeitherDeletesDataNorDescriptor) {
    std::array<BYTE, 8> buffer = { 0, 1, 2, 3, 4, 5, 6, 7 }; // buffer[i] = i

    SAFEARRAY * sa_ptr = nullptr;
    {
        // create an array
        CComSafeArray<BYTE> arr;
        arr.Attach(CreateWeakSafeArray(buffer, FADF_AUTO)); // tag data NOT to be cleared nor deleted at destruction

        sa_ptr = arr.m_psa;
        // SafeArrayDestroy in "arr" dtor does not delete data or descriptor
    }

    // verify that array is NOT cleared and still accessible
    ASSERT_EQ(sa_ptr->cDims, 1);
    ASSERT_EQ(sa_ptr->cbElements, 1u);
    for (size_t i = 0; i < buffer.size(); ++i)
        ASSERT_EQ(reinterpret_cast<BYTE*>(sa_ptr->pvData)[i], i);
    
    // manually delete SAFEARRAY descriptor
    CHECK(SafeArrayDestroyDescriptor(sa_ptr));
}


TEST(SafeArrayTests, VerifyThat_FADF_EMBEDDED_NeitherDeletesDataNorDescriptor) {
    std::array<BYTE, 8> buffer = { 0, 1, 2, 3, 4, 5, 6, 7 }; // buffer[i] = i

    SAFEARRAY * sa_ptr = nullptr;
    {
        // create an array
        CComSafeArray<BYTE> arr;
        arr.Attach(CreateWeakSafeArray(buffer, FADF_EMBEDDED)); // tag data NOT to be cleared nor deleted at destruction

        sa_ptr = arr.m_psa;
        // SafeArrayDestroy in "arr" dtor does not delete data or descriptor
    }

    // verify that array is NOT cleared and still accessible
    ASSERT_EQ(sa_ptr->cDims, 1);
    ASSERT_EQ(sa_ptr->cbElements, 1u);
    for (size_t i = 0; i < buffer.size(); ++i)
        ASSERT_EQ(reinterpret_cast<BYTE*>(sa_ptr->pvData)[i], i);

    // manually delete SAFEARRAY descriptor
    CHECK(SafeArrayDestroyDescriptor(sa_ptr));
}


TEST(SafeArrayTests, TestSafeArray1d) {
    // create an empty array
    CComSafeArray<BYTE> arr((ULONG)0);

    {
        // append data to the array
        BYTE data[] = {0, 1, 2, 3, 4};
        CHECK(arr.Add(ARRAYSIZE(data), data));
    }

    // verify packed data storage
    {
        BYTE * data = nullptr;
        CHECK(SafeArrayAccessData(arr.m_psa, reinterpret_cast<void**>(&data)));
        for (size_t i = 0; i < arr.GetCount(0); ++i)
            ASSERT_EQ(data[i], i);
    
        CHECK(SafeArrayUnaccessData(arr.m_psa));
    }
}


TEST(SafeArrayTests, TestSafeArray2d) {
    // create a 4x8 byte array
    CComSafeArrayBound bounds[] = {4, 8};
    CComSafeArray<BYTE> arr = CreateSafeArray(bounds);

    arr.m_psa->fFeatures |= FADF_FIXEDSIZE; // non-resizable

    // fill array with values in increasing order
    BYTE val = 0;
    for (ULONG j = 0; j < bounds[1].GetCount(); ++j) {
        for (ULONG i = 0; i < bounds[0].GetCount(); ++i) {
            long idx[2] = {static_cast<long>(i), static_cast<long>(j)};
            CHECK(arr.MultiDimSetAt(idx, val));
            val++;
        }
    }

    // verify packed data storage
    {
        BYTE * data = nullptr;
        CHECK(SafeArrayAccessData(arr.m_psa, reinterpret_cast<void**>(&data)));
        for (size_t i = 0; i < bounds[0].GetCount()*bounds[1].GetCount(); ++i)
            ASSERT_EQ(data[i], i);
        
        CHECK(SafeArrayUnaccessData(arr.m_psa));
    }
}
#endif

TEST(SafeArrayTests, LargeSafeArrayTest) {
    {
        // try allocating a 1GB safearray (verify that it doesn't fail)
        ULONG array_size = 1024 * 1024 * 1024;
        CComSafeArray<BYTE> tmp(array_size);
    }
}

TEST(SafeArrayTests, SafeArrayElementAccess) {
    // test CComSafeArray element access
    CComSafeArray<int> buffer(3);

    buffer.GetAt(0) = 0xDE;
    ASSERT_EQ(buffer[0], 0xDE);

    buffer.SetAt(1, 0xAD);
    ASSERT_EQ(buffer[1], 0xAD);

    buffer[2] = 0xBE;
    ASSERT_EQ(buffer.GetAt(2), 0xBE);
    
    buffer.Add(-41);
    ASSERT_EQ(buffer.GetAt(3), -41);
}

TEST(SafeArrayTests, TestEmptySafeArray) {
    {
        // unininitialized SafeArray
        CComSafeArray<BYTE> zero;
        ASSERT_EQ((SAFEARRAY*)zero, nullptr);

        // verify that adding will initialize array
        zero.Add(41);
        ASSERT_NE((SAFEARRAY*)zero, nullptr);
        ASSERT_EQ(zero.GetCount(), 1u);
    }

    {
        // zero-sized SafeArray
        CComSafeArray<BYTE> empty((ULONG)0);
        ASSERT_NE((SAFEARRAY*)empty, nullptr);

        // verify that adding will grow array
        empty.Add(41);
        ASSERT_EQ(empty.GetCount(), 1u);
    }
}

TEST(CComSafeArrayTest, TestSetAtInDifferentConditions) {
    struct MockUnknown : public IUnknown {
        ULONG AddRef() override { return ++ref_count; }
        ULONG Release() override { 
            ULONG result = --ref_count;
            if (result == 0) delete this;
            return result;
        }
        HRESULT QueryInterface(const GUID&, void**) override { return E_NOINTERFACE; }
        std::atomic<ULONG> ref_count{0};
    };

    ATL::CComPtr<IUnknown> obj1 = new MockUnknown();
    ATL::CComPtr<IUnknown> obj2 = new MockUnknown();
    ATL::CComSafeArray<IUnknown*> arr(3);

    //Pass valid objects with valid index
    EXPECT_EQ(S_OK, arr.SetAt(0, obj1));
    EXPECT_EQ(S_OK, arr.SetAt(1, obj2));

    //Pass invalid object with valid index
#ifndef NDEBUG
  #ifndef _WIN32
    EXPECT_DEATH(arr.SetAt(2, nullptr), ""); // triggers assertion failure (internally handled on Windows, unless a debugger is attached)
  #endif
#endif

    //Check that the valid objects are returned from getat
    EXPECT_EQ(arr.GetAt(0), obj1.p);
    EXPECT_EQ(arr.GetAt(1), obj2.p);

    //The one with invalid object should have nullptr
    EXPECT_EQ(arr.GetAt(2), nullptr);

    EXPECT_EQ(3u, arr.GetCount());

    //Try to call SetAt with valid object but invalid index
#ifndef NDEBUG
  #ifndef _WIN32
    EXPECT_DEATH(arr.SetAt(5, obj1), ""); // triggers assertion failure (internally handled on Windows, unless a debugger is attached)
    EXPECT_DEATH(arr.SetAt(-1, obj1), ""); // triggers assertion failure (internally handled on Windows, unless a debugger is attached)
  #endif
#endif
}

TEST(CComSafeArrayTest, TestConvertToSafeArray) {
    std::vector<double> vals = {2.0, 3.0, 4.0};
    {
        // direct assigmnent
        CComSafeArray<double> sa_vals = ConvertToSafeArray(vals.data(), vals.size());
        EXPECT_EQ(sa_vals.GetCount(), 3u);
    }
    {
        // copy-assignment
        CComSafeArray<double> sa_vals;
        sa_vals = ConvertToSafeArray(vals.data(), vals.size());
        EXPECT_EQ(sa_vals.GetCount(), 3u);
    }
}
