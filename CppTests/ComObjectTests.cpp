#include <AppAPI/ComSupport.hpp>
#include <gtest/gtest.h>


struct DECLSPEC_UUID("5E6F7081-9243-4354-6576-8798A9BACBDC")
ICounted : public IUnknown {
    virtual HRESULT Value (int* out) = 0;
};

/** An interface nothing implements, for the E_NOINTERFACE paths. */
struct DECLSPEC_UUID("6F708192-A354-4465-7687-98A9BACBDCED")
IAbsent : public IUnknown {
    virtual HRESULT Unused () = 0;
};

#ifndef _WIN32
static constexpr GUID IID_ICounted = {0x5E6F7081,0x9243,0x4354,{0x65,0x76,0x87,0x98,0xA9,0xBA,0xCB,0xDC}};
static constexpr GUID IID_IAbsent  = {0x6F708192,0xA354,0x4465,{0x76,0x87,0x98,0xA9,0xBA,0xCB,0xDC,0xED}};
DEFINE_UUIDOF(ICounted);
DEFINE_UUIDOF(IAbsent);
#endif

static int g_counted_destroyed = 0;

class Counted : public CComObjectRootEx<CComMultiThreadModel>, public ICounted {
public:
    BEGIN_COM_MAP(Counted)
        COM_INTERFACE_ENTRY(ICounted)
    END_COM_MAP()

    ~Counted () {
        ++g_counted_destroyed;
    }

    HRESULT Value (int* out) override {
        if (!out)
            return E_POINTER;
        *out = 7;
        return S_OK;
    }
};


/** Current reference count, without changing it. AddRef/Release return the new
    value, and the member holding it is named differently by ATL and MiniCOM. */
static ULONG RefCount (IUnknown* obj) {
    ULONG count = obj->AddRef();
    obj->Release();
    return count - 1;
}

TEST(ComObjectTests, CreateInstanceStartsAtZeroAndDestroysOnLastRelease) {
    g_counted_destroyed = 0;
    CComObject<Counted>* obj = nullptr;
    ASSERT_EQ(CComObject<Counted>::CreateInstance(&obj), S_OK);
    ASSERT_TRUE(obj);
    EXPECT_EQ(obj->AddRef(), 1u);   // CreateInstance hands it back unreferenced
    ICounted* counted = nullptr;
    ASSERT_EQ(obj->QueryInterface(__uuidof(ICounted), (void**)&counted), S_OK);
    EXPECT_EQ(RefCount(obj), 2u);
    int value = 0;
    EXPECT_EQ(counted->Value(&value), S_OK);
    EXPECT_EQ(value, 7);

    IUnknown* unknown = nullptr;
    ASSERT_EQ(obj->QueryInterface(__uuidof(IUnknown), (void**)&unknown), S_OK);
    EXPECT_EQ(RefCount(obj), 3u);

    void* absent = (void*)0x1;
    EXPECT_EQ(obj->QueryInterface(__uuidof(IAbsent), &absent), E_NOINTERFACE);
    EXPECT_EQ(absent, nullptr);

    EXPECT_EQ(unknown->Release(), 2u);
    EXPECT_EQ(counted->Release(), 1u);
    EXPECT_EQ(g_counted_destroyed, 0);
    EXPECT_EQ(obj->Release(), 0u);
    EXPECT_EQ(g_counted_destroyed, 1);
}

TEST(ComObjectTests, CComPtrCopyToAndIsEqualObject) {
    g_counted_destroyed = 0;
    CComObject<Counted>* obj = nullptr;
    ASSERT_EQ(CComObject<Counted>::CreateInstance(&obj), S_OK);
    obj->AddRef();
    EXPECT_EQ(RefCount(obj), 1u);

    {
        CComPtr<ICounted> p(static_cast<ICounted*>(obj));
        EXPECT_EQ(RefCount(obj), 2u);
        EXPECT_TRUE(p.IsEqualObject(static_cast<IUnknown*>(obj)));

        CComPtr<ICounted> copied;
        EXPECT_EQ(p.CopyTo(&copied), S_OK);
        EXPECT_EQ(RefCount(obj), 3u);
    }
    EXPECT_EQ(RefCount(obj), 1u);

    EXPECT_EQ(g_counted_destroyed, 0);
    EXPECT_EQ(obj->Release(), 0u);
    EXPECT_EQ(g_counted_destroyed, 1);
}
