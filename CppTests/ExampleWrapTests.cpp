#include <AppAPI/ComSupport.hpp>
#include "ExampleWrap.tlh"
#include <gtest/gtest.h>


/** COM class implementing the TLH-generated interfaces. */
class WrapCalculator :
    public CComObjectRootEx<CComMultiThreadModel>,
    public CComCoClass<WrapCalculator>,
    public ExampleLib::ICalcExt, public ExampleLib::ICalc2 {
public:
    HRESULT raw_GetValue (/*out*/int * value) override {
        *value = 42;
        return S_OK;
    }

    HRESULT raw_Add (int a, int b, /*out*/int * result) override {
        *result = a + b;
        return S_OK;
    }

    HRESULT raw_SetCallback (ExampleLib::ICalcCb * cb) override {
        m_callback = cb;
        return S_OK;
    }

    HRESULT raw_GetValue2 (/*out*/int * value) override {
        *value = 43;
        return S_OK;
    }

    ExampleLib::ICalcCbPtr m_callback;

    BEGIN_COM_MAP(WrapCalculator)
        COM_INTERFACE_ENTRY(ExampleLib::ICalc)
        COM_INTERFACE_ENTRY(ExampleLib::ICalcExt)
        COM_INTERFACE_ENTRY(ExampleLib::ICalc2)
    END_COM_MAP()
};


TEST(ExampleWrapTests, WrapperMethods) {
    CComPtr<WrapCalculator> calc = CreateLocalInstance<WrapCalculator>();

    ExampleLib::ICalcExtPtr ext(static_cast<ExampleLib::ICalcExt*>(calc));
    EXPECT_EQ(ext->GetValue(), 42); // wrapper method returns the value directly
    EXPECT_EQ(ext->Add(1, 2), 3);

    ExampleLib::ICalc2Ptr two(ext); // implicit QueryInterface
    EXPECT_EQ(two->GetValue2(), 43);
}
