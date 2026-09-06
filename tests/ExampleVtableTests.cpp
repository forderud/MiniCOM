#define CINTERFACE // access C language vtable definitions instead of C++ interfaces
#ifndef WIN32
  #include "NonWindows.hpp"
#endif
#include "Example.h"
#include <gtest/gtest.h>


template <class IUNKNOWN>
static void TestAddRefRelease(IUNKNOWN* ptr) {
    ULONG ref1 = ptr->lpVtbl->AddRef(ptr);
    EXPECT_GT(ref1, 1); // ref-count >1
    ULONG ref2 = ptr->lpVtbl->Release(ptr);
    EXPECT_GT(ref2, 0); // ref-count >0
    EXPECT_EQ(ref2, ref1 - 1);
}

template <class INTERFACE>
static void TestQueryInterface(INTERFACE* ptr) {
    // try to cast to IUnknown (increases the ref-count)
    IUnknown* unknown = nullptr;
    HRESULT hr = ptr->lpVtbl->QueryInterface(ptr, IID_IUnknown, (void**)&unknown);
    EXPECT_EQ(hr, S_OK);

    // clean up reference
    unknown->lpVtbl->Release(unknown);
}


void TestICalcVtable(ICalc* ptr) {
#ifdef _WIN32
    TestAddRefRelease(ptr);
    TestQueryInterface(ptr);

    int val = 0;
    HRESULT hr = ptr->lpVtbl->GetValue(ptr, &val);
    EXPECT_EQ(hr, S_OK);
    EXPECT_EQ(val, 42);
#endif
}

void TestICalcExtVtable(ICalcExt* ptr) {
#ifdef _WIN32
    TestAddRefRelease(ptr);
    TestQueryInterface(ptr);

    int val = 0;
    HRESULT hr = ptr->lpVtbl->GetValue(ptr, &val);
    EXPECT_EQ(hr, S_OK);
    EXPECT_EQ(val, 42);

    hr = ptr->lpVtbl->Add(ptr, 1, 2, &val);
    EXPECT_EQ(hr, S_OK);
    EXPECT_EQ(val, 3);
#endif
}

void TestICalc2Vtable(ICalc2* ptr) {
#ifdef _WIN32
    TestAddRefRelease(ptr);
    TestQueryInterface(ptr);

    int val = 0;
    HRESULT hr = ptr->lpVtbl->GetValue2(ptr, &val);
    EXPECT_EQ(hr, S_OK);
    EXPECT_EQ(val, 43);
#endif
}
