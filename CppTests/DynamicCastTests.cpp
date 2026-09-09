#include <gtest/gtest.h>
#include <AppAPI/ComSupport.hpp>

struct DECLSPEC_UUID("672147B6-F19F-4F9D-A647-382F27756B78")
IMyLogger : public IUnknown {
    virtual HRESULT Log(/*in*/BSTR msg) = 0;
};
#ifndef _WIN32
static constexpr GUID IID_IMyLogger = { 0x672147B6, 0xF19F, 0x4F9D,{0xA6,0x47,0x38,0x2F,0x27,0x75,0x6B,0x78} };
DEFINE_UUIDOF(IMyLogger);
#endif


/** COM class with vtable. */
class LoggerVtable :
    public CComObjectRootEx<CComMultiThreadModel>, // also compatible with STA
    public CComCoClass<LoggerVtable>, // no CLSID needed
    public IMyLogger {
public:
    HRESULT Log(BSTR /*msg*/) override {
        return E_NOTIMPL;
    }

    BEGIN_COM_MAP(LoggerVtable)
        COM_INTERFACE_ENTRY(IMyLogger)
    END_COM_MAP()
};

/** COM class without vtable. */
class ATL_NO_VTABLE LoggerNoVtable :
    public CComObjectRootEx<CComMultiThreadModel>, // also compatible with STA
    public CComCoClass<LoggerNoVtable>, // no CLSID needed
    public IMyLogger {
public:
    HRESULT Log(BSTR /*msg*/) override {
        return E_NOTIMPL;
    }

    BEGIN_COM_MAP(LoggerNoVtable)
        COM_INTERFACE_ENTRY(IMyLogger)
    END_COM_MAP()
};


TEST(CastTests, verify_dynamic_cast_works_by_default) {
    // create COM object with vtable
    CComPtr<IMyLogger> logger;
    logger = CreateLocalInstance<LoggerVtable>();
    // get interface pointer
    IMyLogger* ptr = logger.p;

    {
        // try dynamic_cast to correct underlying C++ class
        auto* obj = dynamic_cast<LoggerVtable*>(ptr); // should succeed
        ASSERT_NE(obj, nullptr);
    }
    {
        // try dynamic_cast to wrond underlying C++ class
        auto* obj = dynamic_cast<LoggerNoVtable*>(ptr); // should fail
        ASSERT_EQ(obj, nullptr);
    }
}

TEST(CastTests, verify_dynamic_cast_works_without_vtable) {
    // create COM object without vtable
    CComPtr<IMyLogger> logger;
    logger = CreateLocalInstance<LoggerNoVtable>();
    // get interface pointer
    IMyLogger* ptr = logger.p;
    {
        // try dynamic_cast to correct underlying C++ class
        auto* obj = dynamic_cast<LoggerNoVtable*>(ptr); // should succeed
        ASSERT_NE(obj, nullptr);
    }
    {
        // try dynamic_cast to wrond underlying C++ class
        auto* obj = dynamic_cast<LoggerVtable*>(ptr); // should fail
        ASSERT_EQ(obj, nullptr);
    }
}
