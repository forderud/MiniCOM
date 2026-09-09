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


class SimpleLogger :
    public CComObjectRootEx<CComMultiThreadModel>, // also compatible with STA
    public CComCoClass<SimpleLogger>, // no CLSID needed
    public IMyLogger {
public:
    HRESULT Log(BSTR msg) override {
        if (!msg)
            return E_INVALIDARG;

        std::wcout << L"Log: " << msg << std::endl;
        return S_OK;
    }

    BEGIN_COM_MAP(SimpleLogger)
        COM_INTERFACE_ENTRY(IMyLogger)
    END_COM_MAP()
};


/** Templatized COM class. */
template <class T>
class TemplatedLogger :
    public CComObjectRootEx<CComMultiThreadModel>, // also compatible with STA
    public CComCoClass<TemplatedLogger<T>>, // no CLSID needed
    public IMyLogger {
public:
    HRESULT Log(BSTR msg) override {
        if (!msg)
            return E_INVALIDARG;

        std::wcout << L"Log: " << msg << std::endl;
        return S_OK;
    }

    BEGIN_COM_MAP(TemplatedLogger)
        COM_INTERFACE_ENTRY(IMyLogger)
    END_COM_MAP()

private:
    T m_value = 41;
};

template<>
HRESULT TemplatedLogger<float>::Log(BSTR msg) {
    if (!msg)
        return E_INVALIDARG;

    std::wcout << L"Log: " << msg << L" " << m_value << std::endl;
    return S_OK;
};


template <>
class TemplatedLogger<int> : public CComObjectRootEx<CComMultiThreadModel>, // also compatible with STA
    public CComCoClass<TemplatedLogger<int>>, // no CLSID needed
    public IMyLogger {
public:
    HRESULT Log(BSTR msg) override {
        if (!msg)
            return E_INVALIDARG;

        std::wcout << L"Log: " << msg << L" <int>" << std::endl;
        return S_OK;
    }

    BEGIN_COM_MAP(TemplatedLogger)
        COM_INTERFACE_ENTRY(IMyLogger)
    END_COM_MAP()
};



TEST(ComClassTests, TestSimpleLogger) {
    CComPtr<IMyLogger> logger;
    logger = CreateLocalInstance<SimpleLogger>();
    logger->Log(CComBSTR(L"Hi SimpleLogger"));
}

TEST(ComClassTests, TestTemplatedLogger) {
    CComPtr<IMyLogger> logger;
    logger = CreateLocalInstance<TemplatedLogger<float>>();
    logger->Log(CComBSTR(L"Hi TemplatedLogger"));

    logger = CreateLocalInstance<TemplatedLogger<int>>();
    logger->Log(CComBSTR(L"Hi TemplatedLogger"));
}

TEST(ComClassTests, QueryInterfaceRejectsNullOutPointer) {
    CComPtr<IMyLogger> logger;
    logger = CreateLocalInstance<SimpleLogger>();

    EXPECT_EQ(logger->QueryInterface(__uuidof(IMyLogger), nullptr), E_POINTER);
}
