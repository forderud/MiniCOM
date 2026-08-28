#pragma once
#include <AppAPI/ComSupport.hpp>


struct DECLSPEC_UUID("672147B6-F19F-4F9D-A647-382F27756B78")
IMyLogger :
#ifndef _WIN32
    virtual
#endif
    public IUnknown {
    virtual HRESULT Log(/*in*/BSTR msg) = 0;
};
#ifndef _WIN32
static constexpr GUID IID_IMyLogger = { 0x672147B6, 0xF19F, 0x4F9D,{0xA6,0x47,0x38,0x2F,0x27,0x75,0x6B,0x78} };
DEFINE_UUIDOF(IMyLogger);
#endif

struct DECLSPEC_UUID("A768B57D-63A0-49FE-98DD-43CF98D2D2A7")
IMyMessage :
#ifndef _WIN32
    virtual
#endif
    public IUnknown {
    virtual HRESULT ShowMessage(/*in*/BSTR msg) = 0;
};
#ifndef _WIN32
static constexpr GUID IID_IMyMessage = { 0xA768B57D, 0x63A0, 0x49FE,{0x98,0xDD,0x43,0xCF,0x98,0xD2,0xD2,0xA7} };
DEFINE_UUIDOF(IMyMessage);
#endif


static unsigned int GetRefCount(IUnknown& obj) {
    obj.AddRef();
    return obj.Release();
}

/** Inner logger class for testing of COM aggregation. */
class InnerLogger :
    public CComObjectRootEx<CComMultiThreadModel>, // also compatible with STA
    public CComCoClass<InnerLogger>, // no CLSID needed
    public IMyLogger {
public:
    InnerLogger() {
        std::wcout << L"InnerLogger ctor" << std::endl;
        s_obj_count++;
    }
    ~InnerLogger() {
        std::wcout << L"InnerLogger dtor" << std::endl;
        s_obj_count--;
    }

    HRESULT Log(BSTR msg) override {
        if (!msg)
            return E_INVALIDARG;

        std::wcout << L"Log: " << msg << std::endl;
        return S_OK;
    }

    BEGIN_COM_MAP(InnerLogger)
        COM_INTERFACE_ENTRY(IMyLogger)
    END_COM_MAP()

    static inline ULONG s_obj_count = 0;
};
