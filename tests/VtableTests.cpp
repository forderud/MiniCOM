#include <AppAPI/ComSupport.hpp>
#include "Example.h"
#include <gtest/gtest.h>


/** COM class for checking the generated vtable structs against the vtables
    built by the C++ compiler. Implements two interfaces, so that ICalc2 is
    reached through a secondary base. */
class VtableCalculator :
    public CComObjectRootEx<CComMultiThreadModel>, // also compatible with STA
    public CComCoClass<VtableCalculator>, // no CLSID needed
    public ICalcExt, public ICalc2 {
public:
    HRESULT GetValue (/*out*/int * value) override {
        *value = 42;
        return S_OK;
    }

    HRESULT Add (int a, int b, /*out*/int * result) override {
        *result = a + b;
        return S_OK;
    }

    HRESULT GetValue2 (/*out*/int * value) override {
        *value = 43;
        return S_OK;
    }

    BEGIN_COM_MAP(VtableCalculator)
        COM_INTERFACE_ENTRY(ICalc)
        COM_INTERFACE_ENTRY(ICalcExt)
        COM_INTERFACE_ENTRY(ICalc2)
    END_COM_MAP()
};


/** Read the vtable pointer out of an object, which is all a caller outside C++
    has to go on: an interface pointer, and the layout it expects to find. */
template <class VTBL, class T>
static const VTBL* VtableOf (T * obj) {
    return *reinterpret_cast<const VTBL* const*>(obj);
}


TEST(VtableTests, TestMethodCallsThroughVtable) {
    CComPtr<VtableCalculator> calc = CreateLocalInstance<VtableCalculator>();

    CComPtr<ICalcExt> ptr;
    HRESULT hr = calc.QueryInterface(&ptr);
    EXPECT_EQ(hr, S_OK);

    const ICalcExtVtbl* vtbl = VtableOf<ICalcExtVtbl>(ptr.p);

    int val = 0;
    hr = vtbl->GetValue(ptr, &val); // inherited from ICalc
    EXPECT_EQ(hr, S_OK);
    EXPECT_EQ(val, 42);

    hr = vtbl->Add(ptr, 1, 2, &val);
    EXPECT_EQ(hr, S_OK);
    EXPECT_EQ(val, 3);
}

TEST(VtableTests, TestIUnknownSlotsComeFirst) {
    CComPtr<VtableCalculator> calc = CreateLocalInstance<VtableCalculator>();

    CComPtr<ICalcExt> ptr;
    HRESULT hr = calc.QueryInterface(&ptr);
    EXPECT_EQ(hr, S_OK);

    const ICalcExtVtbl* vtbl = VtableOf<ICalcExtVtbl>(ptr.p);

    ULONG ref = vtbl->AddRef(ptr);
    EXPECT_EQ(vtbl->Release(ptr), ref - 1);

    CComPtr<ICalc> base;
    hr = vtbl->QueryInterface(ptr, __uuidof(ICalc), (void**)&base);
    EXPECT_EQ(hr, S_OK);
    ASSERT_TRUE(base);

    // the base interface lays out the same way, minus the derived methods
    int val = 0;
    hr = VtableOf<ICalcVtbl>(base.p)->GetValue(base, &val);
    EXPECT_EQ(hr, S_OK);
    EXPECT_EQ(val, 42);
}

TEST(VtableTests, TestSecondaryInterface) {
    // ICalc2 is a second base, so it sits at a different offset with a vtable
    // of its own, and the interface pointer is adjusted to match.
    CComPtr<VtableCalculator> calc = CreateLocalInstance<VtableCalculator>();

    CComPtr<ICalc2> ptr;
    HRESULT hr = calc.QueryInterface(&ptr);
    EXPECT_EQ(hr, S_OK);

    const ICalc2Vtbl* vtbl = VtableOf<ICalc2Vtbl>(ptr.p);

    int val = 0;
    hr = vtbl->GetValue2(ptr, &val);
    EXPECT_EQ(hr, S_OK);
    EXPECT_EQ(val, 43);

    CComPtr<ICalc> back;
    hr = vtbl->QueryInterface(ptr, __uuidof(ICalc), (void**)&back);
    EXPECT_EQ(hr, S_OK);

    hr = VtableOf<ICalcVtbl>(back.p)->GetValue(back, &val);
    EXPECT_EQ(hr, S_OK);
    EXPECT_EQ(val, 42);
}
