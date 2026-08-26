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
_COM_SMARTPTR_TYPEDEF(IExample, __uuidof(IExample));

static constexpr GUID IID_IUnrelated = {0x11112222, 0x3333, 0x4444, {5, 6, 7, 8, 9, 10, 11, 12}};

struct IUnrelated : virtual public IUnknown {
    virtual HRESULT Unused () = 0;
};
DEFINE_UUIDOF(IUnrelated)
_COM_SMARTPTR_TYPEDEF(IUnrelated, __uuidof(IUnrelated));

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


/** A class whose construction fails, to check the error reaches the caller. */
static constexpr GUID CLSID_Unbuildable = {0x3C4D5E6F, 0x7081, 0x4292, {0xA3, 0xB4, 0xC5, 0xD6, 0xE7, 0xF8, 0x09, 0x1B}};

class Unbuildable : public CComObjectRootEx<CComMultiThreadModel>, public IExample {
public:
    BEGIN_COM_MAP(Unbuildable)
        COM_INTERFACE_ENTRY(IExample)
    END_COM_MAP()

    HRESULT FinalConstruct () {
        return E_OUTOFMEMORY;
    }

    HRESULT GetValue (int* out) override {
        (void)out;
        return E_UNEXPECTED; // never reached: the object is never handed out
    }
};

OBJECT_ENTRY_AUTO(CLSID_Unbuildable, Unbuildable)

static constexpr GUID CLSID_Missing = {0xDEADBEEF, 0x0000, 0x0000, {0, 0, 0, 0, 0, 0, 0, 0}};


void TestCreateInstanceErrors () {
    printf("an unknown class reaches C++ callers as an error, not an abort...\n");
    {
        IExamplePtr p;
        assert(p.CreateInstance(CLSID_Missing) == REGDB_E_CLASSNOTREG);
        assert(!p);
        assert(p.CreateInstance(L"NoSuchClass") == CO_E_CLASSSTRING);
        assert(!p);
    }
    {
        ATL::CComPtr<IExample> p;
        assert(p.CoCreateInstance(CLSID_Missing) == REGDB_E_CLASSNOTREG);
        assert(!p);
        assert(p.CoCreateInstance(L"NoSuchClass") == CO_E_CLASSSTRING);
        assert(!p);
    }

    printf("...and the happy path is unchanged...\n");
    {
        IExamplePtr p;
        assert(p.CreateInstance(CLSID_Example) == S_OK);
        int value = 0;
        assert(p->GetValue(&value) == S_OK && value == 42);

        ATL::CComPtr<IExample> q;
        assert(q.CoCreateInstance(L"Example") == S_OK);
        assert(q->GetValue(&value) == S_OK && value == 42);
    }

    printf("...and a failing factory reports its own error...\n");
    {
        IExamplePtr p;
        assert(p.CreateInstance(CLSID_Unbuildable) == E_OUTOFMEMORY);
        assert(!p);

        void* obj = (void*)0x1;
        assert(CoCreateInstance(CLSID_Unbuildable, nullptr, CLSCTX_INPROC_SERVER,
                                __uuidof(IExample), &obj) == E_OUTOFMEMORY);
        assert(obj == nullptr);
    }

    printf("...an interface the class does not support still gives E_NOINTERFACE...\n");
    {
        IUnrelatedPtr p;
        assert(p.CreateInstance(CLSID_Example) == E_NOINTERFACE);
        assert(!p);
    }
}


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
    void* obj = (void*)0x1;
    assert(CoCreateInstance(CLSID_Missing, nullptr, CLSCTX_INPROC_SERVER,
                            __uuidof(IUnknown), &obj) == REGDB_E_CLASSNOTREG);
    assert(obj == nullptr);   // cleared even on failure, as ole32 does

    printf("an unsupported interface gives E_NOINTERFACE...\n");
    assert(CoCreateInstance(CLSID_Example, nullptr, CLSCTX_INPROC_SERVER,
                            IID_IUnrelated, &obj) == E_NOINTERFACE);

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
    TestCreateInstanceErrors();
}
