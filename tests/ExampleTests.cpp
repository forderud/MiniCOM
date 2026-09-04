#include <AppAPI/ComSupport.hpp>
#include "Example.h"
#include <gtest/gtest.h>


class Calculator : 
    public CComObjectRootEx<CComMultiThreadModel>, // also compatible with STA
    public CComCoClass<Calculator>, // no CLSID needed
    public ICalcExt, public ICalc2 {
public:
    Calculator() {
    }

    ~Calculator() {
    }

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

    BEGIN_COM_MAP(Calculator)
        COM_INTERFACE_ENTRY(ICalc)
        COM_INTERFACE_ENTRY(ICalcExt)
        COM_INTERFACE_ENTRY(ICalc2)
    END_COM_MAP()
};


TEST(ExampleTests, TestConstructionAndMethods) {
    CComPtr<Calculator> calc = CreateLocalInstance<Calculator>();
    
    {
        CComPtr<ICalc> ptr;
        HRESULT hr = calc.QueryInterface(&ptr);
        EXPECT_EQ(hr, S_OK);
        int val = 0;
        hr = ptr->GetValue(&val);
        EXPECT_EQ(hr, S_OK);
        EXPECT_EQ(val, 42);
    }
    
    {
        CComPtr<ICalcExt> ptr;
        HRESULT hr = calc.QueryInterface(&ptr);
        EXPECT_EQ(hr, S_OK);
        int val = 0;
        hr = ptr->Add(1, 2, &val);
        EXPECT_EQ(hr, S_OK);
        EXPECT_EQ(val, 3);
    }

    {
        CComPtr<ICalc2> ptr;
        HRESULT hr = calc.QueryInterface(&ptr);
        EXPECT_EQ(hr, S_OK);
        int val = 0;
        hr = ptr->GetValue2(&val);
        EXPECT_EQ(hr, S_OK);
        EXPECT_EQ(val, 43);
    }
}
