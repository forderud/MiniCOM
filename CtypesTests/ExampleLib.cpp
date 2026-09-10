/** Shared library with the example COM class, so that a client outside C++ has
    something to activate. Only the CLSID and the IDL description are needed to
    reach it, the same as for any other COM server. */

#include <AppAPI/ComSupport.hpp>
#include "Example.h"


/** Implementation of the Calculator coclass. */
class CalculatorImpl :
    public CComObjectRootEx<CComMultiThreadModel>, // also compatible with STA
    public CComCoClass<CalculatorImpl>,
    public ICalcExt, public ICalc2 {
public:
    HRESULT GetValue (/*out*/int * value) override {
        if (!value)
            return E_INVALIDARG;

        *value = 42;
        return S_OK;
    }

    HRESULT Add (int a, int b, /*out*/int * result) override {
        if (!result)
            return E_INVALIDARG;

        *result = a + b;
        return S_OK;
    }

    HRESULT SetCallback (ICalcCb * cb) override {
        if (!cb)
            return E_INVALIDARG;

        m_callback = cb;
        return S_OK;
    }

    HRESULT GetValue2 (/*out*/int * value) override {
        if (!value)
            return E_INVALIDARG;

        *value = 43;

        if (m_callback)
            m_callback->Message(CComBSTR(L"GetValue2 called"));

        return S_OK;
    }

    BEGIN_COM_MAP(CalculatorImpl)
        COM_INTERFACE_ENTRY(ICalc)
        COM_INTERFACE_ENTRY(ICalcExt)
        COM_INTERFACE_ENTRY(ICalc2)
    END_COM_MAP()

private:
    CComPtr<ICalcCb> m_callback;
};
OBJECT_ENTRY_AUTO(CLSID_Calculator, CalculatorImpl)
