#include <AppAPI/ComSupport.hpp>
#include <gtest/gtest.h>


static constexpr GUID IID_IExample = {0x2B3C4D5E, 0x6F70, 0x4181, {0x92, 0x93, 0xA4, 0xB5, 0xC6, 0xD7, 0xE8, 0xF9}};
static constexpr GUID IID_IMissing = {0x11112222, 0x3333, 0x4444, {5, 6, 7, 8, 9, 10, 11, 12}};
static constexpr GUID CLSID_Example = {0x3C4D5E6F, 0x7081, 0x4292, {0xA3, 0xB4, 0xC5, 0xD6, 0xE7, 0xF8, 0x09, 0x1A}};

struct IExample : public IUnknown {
    virtual HRESULT GetValue (int* out) = 0;
};
DEFINE_UUIDOF(IExample)

struct IMissing : public IUnknown {
    virtual HRESULT Unused () = 0;
};
DEFINE_UUIDOF(IMissing)

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


TEST(ComCreateTests, CoCreateInstanceReturnsRequestedInterface) {
    CComPtr<IExample> example;
    ASSERT_EQ(example.CoCreateInstance(CLSID_Example), S_OK);
    ASSERT_TRUE(example);

    int value = 0;
    EXPECT_EQ(example->GetValue(&value), S_OK);
    EXPECT_EQ(value, 42);
}

TEST(ComCreateTests, CoCreateInstanceByProgID) {
    CComPtr<IExample> example;
    ASSERT_EQ(example.CoCreateInstance(L"Example"), S_OK);
    ASSERT_TRUE(example);

    int value = 0;
    EXPECT_EQ(example->GetValue(&value), S_OK);
    EXPECT_EQ(value, 42);
}

TEST(ComCreateTests, UnregisteredClassAndInterfaceErrors) {
    GUID unknown_clsid = {0xDEADBEEF, 0x0000, 0x0000, {0, 0, 0, 0, 0, 0, 0, 0}};
    CComPtr<IUnknown> obj;
    EXPECT_EQ(obj.CoCreateInstance(unknown_clsid), REGDB_E_CLASSNOTREG);
    EXPECT_FALSE(obj);

    CComPtr<IUnknown> by_name;
    EXPECT_EQ(by_name.CoCreateInstance(L"NoSuchClass"), CO_E_CLASSSTRING);
    EXPECT_FALSE(by_name);

    CComPtr<IMissing> missing;
    EXPECT_EQ(missing.CoCreateInstance(CLSID_Example), E_NOINTERFACE);
    EXPECT_FALSE(missing);
}
