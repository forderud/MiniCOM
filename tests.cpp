#include <cassert>
#include <vector>
#include "NonWindows.hpp"


/** Convert raw array to SafeArray. */
template <class T>
CComSafeArray<T> ConvertToSafeArray (const T * input, size_t element_count) {
    CComSafeArray<T> result(static_cast<ULONG>(element_count));

    if (element_count > 0) {
        T * out_ptr = &result.GetAt(0); // will fail if empty
        for (size_t i = 0; i < element_count; ++i)
            out_ptr[i] = input[i];
    }
    return result;
}


void TestCComSafeArray() {
    std::vector<double> vals = {2.0, 3.0, 4.0};
    {
        printf("direct assigmnent...\n");
        CComSafeArray<double> sa_vals = ConvertToSafeArray(vals.data(), vals.size());
        assert(sa_vals.GetCount() == 3);
    }
    {
        printf("copy-assignment...\n");
        CComSafeArray<double> sa_vals;
        sa_vals = ConvertToSafeArray(vals.data(), vals.size());
        assert(sa_vals.GetCount() == 3);
    }
}

/** A class registered the ordinary way, to create through the ole32-parity API.
    Declared the way IdlParse.py emits interfaces: an IID_ constant, virtual
    inheritance from IUnknown, then DEFINE_UUIDOF. */
static constexpr GUID IID_IExample = {0x2B3C4D5E, 0x6F70, 0x4181, {0x92, 0x93, 0xA4, 0xB5, 0xC6, 0xD7, 0xE8, 0xF9}};

struct IExample : virtual public IUnknown {
    virtual HRESULT GetValue (int* out) = 0;
};
DEFINE_UUIDOF(IExample)

static constexpr GUID CLSID_Example = {0x3C4D5E6F, 0x7081, 0x4292, {0xA3, 0xB4, 0xC5, 0xD6, 0xE7, 0xF8, 0x09, 0x1A}};

class Example : public CComObjectRootEx<CComMultiThreadModel>, public IExample {
public:
    BEGIN_COM_MAP(Example)
        COM_INTERFACE_ENTRY(IExample)
    END_COM_MAP()

    HRESULT GetValue (int* out) override {
        if (!out)
            return E_POINTER;
        *out = 42;
        return S_OK;
    }
};

OBJECT_ENTRY_AUTO(CLSID_Example, Example)


void TestCoCreateInstance () {
    printf("CoCreateInstance returns the interface asked for...\n");
    IExample* example = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_Example, nullptr, CLSCTX_INPROC_SERVER,
                                  __uuidof(IExample), (void**)&example);
    assert(hr == S_OK);
    assert(example != nullptr);
    int value = 0;
    assert(example->GetValue(&value) == S_OK);
    assert(value == 42);
    example->Release();

    printf("...and IUnknown when that is what is asked for...\n");
    IUnknown* unknown = nullptr;
    hr = CoCreateInstance(CLSID_Example, nullptr, CLSCTX_INPROC_SERVER,
                          __uuidof(IUnknown), (void**)&unknown);
    assert(hr == S_OK && unknown != nullptr);
    unknown->Release();

    printf("unregistered CLSID gives REGDB_E_CLASSNOTREG...\n");
    GUID missing = {0xDEADBEEF, 0x0000, 0x0000, {0, 0, 0, 0, 0, 0, 0, 0}};
    void* obj = (void*)0x1;
    assert(CoCreateInstance(missing, nullptr, CLSCTX_INPROC_SERVER,
                            __uuidof(IUnknown), &obj) == REGDB_E_CLASSNOTREG);
    assert(obj == nullptr);   // cleared even on failure, as ole32 does

    printf("an unsupported interface gives E_NOINTERFACE...\n");
    static constexpr GUID IID_Absent = {0x11112222, 0x3333, 0x4444, {5, 6, 7, 8, 9, 10, 11, 12}};
    assert(CoCreateInstance(CLSID_Example, nullptr, CLSCTX_INPROC_SERVER,
                            IID_Absent, &obj) == E_NOINTERFACE);

    printf("CLSIDFromProgID resolves a registered name...\n");
    GUID resolved{};
    assert(CLSIDFromProgID(L"Example", &resolved) == S_OK);
    assert(resolved == CLSID_Example);
    assert(CLSIDFromProgID(L"NoSuchClass", &resolved) == CO_E_CLASSSTRING);

    printf("null arguments are refused...\n");
    assert(CoCreateInstance(CLSID_Example, nullptr, CLSCTX_INPROC_SERVER,
                            __uuidof(IExample), nullptr) == E_POINTER);
    assert(CLSIDFromProgID(nullptr, &resolved) == E_POINTER);
    assert(CLSIDFromProgID(L"Example", nullptr) == E_POINTER);
}


int main() {
    printf("Running tests...\n");
    TestCComSafeArray();
    TestCoCreateInstance();
}
