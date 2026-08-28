#include <AppAPI/ComSupport.hpp>
#include <gtest/gtest.h>

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

struct DECLSPEC_UUID("9F43078A-751F-4A18-93C1-35FA2CC59CB6")
IUnsupported :
#ifndef _WIN32
    virtual
#endif
    public IUnknown {
    virtual HRESULT DoSomething() = 0;
};
#ifndef _WIN32
static constexpr GUID IID_IUnsupported = { 0x9F43078A, 0x751F, 0x4A18,{0x93,0xC1,0x35,0xFA,0x2C,0xC5,0x9C,0xB6} };
DEFINE_UUIDOF(IUnsupported);
#endif

// This define will be removed when switching to generated wrapper-API
_COM_SMARTPTR_TYPEDEF(IMyLogger, __uuidof(IMyLogger));
_COM_SMARTPTR_TYPEDEF(IMyMessage, __uuidof(IMyMessage));
_COM_SMARTPTR_TYPEDEF(IUnsupported, __uuidof(IUnsupported));


class ComObjMock : public IMyMessage, public IMyLogger {
public:
    ComObjMock() {
        ++s_obj_count;
    }

    ~ComObjMock() {
        --s_obj_count;
    }

    HRESULT QueryInterface(const IID& iid, void** ptr) override {
        assert(ptr && !*ptr);

        if (iid == __uuidof(IMyMessage))
            *ptr = static_cast<IMyMessage*>(this);
        else if (iid == __uuidof(IMyLogger))
            *ptr = static_cast<IMyLogger*>(this);
        else if (iid == __uuidof(IUnknown))
            *ptr = static_cast<IUnknown*>(static_cast<IMyMessage*>(this)); // cast through IMyMessage to get IUnknown (convention chosen)
        else
            return E_NOINTERFACE;

        AddRef();
        return S_OK;
    }

    HRESULT Log(BSTR /*message*/) override {
        std::cout << "ComObjMock Log called\n";
        return S_OK;
    }

    ULONG AddRef() override {
        return ++m_ref;
    }

    ULONG Release() override {
        ULONG ref = --m_ref;
        if (!ref)
            delete this;
        return ref;
    }

    HRESULT ShowMessage(BSTR /*msg*/) override {
        return E_NOTIMPL;
    }

    std::atomic<ULONG> m_ref = 0;

    static void LeakCheck() {
        assert(!s_obj_count && "ComObjMock leak detected");
        if (s_obj_count)
            throw std::runtime_error("ComObjMock leak detected");
    }

private:
    static std::atomic<ULONG> s_obj_count;
};

std::atomic<ULONG> ComObjMock::s_obj_count(0);


inline unsigned long GetRefCount(IUnknown& obj) {
    obj.AddRef();
    return obj.Release();
}

class ComPtrTests : public ::testing::Test {
public:
    static void SetUpTestCase() {
        ComObjMock::LeakCheck();
    }

    static void TearDownTestCase() {
        ComObjMock::LeakCheck();
    }
};

TEST_F(ComPtrTests, Test_com_ptr_t_ptr_constructon) {
    // construct from raw ptr
    IUnknownPtr ptr1(new ComObjMock);
    EXPECT_EQ(GetRefCount(*ptr1), 1u);

    // construct from raw ptr
    IMyLoggerPtr ptr2(static_cast<IMyLogger*>(new ComObjMock));
    EXPECT_EQ(GetRefCount(*ptr2), 1u);
}

TEST_F(ComPtrTests, Test_com_ptr_t_assignment1) {
    // ptr-cast construction & assignment
    IUnknownPtr ptr1 = new ComObjMock;

    IUnknownPtr ptr2;
    ptr2 = ptr1;
    EXPECT_TRUE(ptr2);

    EXPECT_EQ(GetRefCount(*ptr1), 2u);
}

TEST_F(ComPtrTests, Test_com_ptr_t_assignment2) {
    // ptr-cast construction & assignment
    IMyLoggerPtr ptr = new ComObjMock;

    // cast that succeed
    IMyLoggerPtr pl;
    pl = ptr;
    EXPECT_TRUE(pl);

    // cast that fail
    IUnsupportedPtr pd = ptr;
    EXPECT_FALSE(pd);

    EXPECT_EQ(GetRefCount(*ptr), 2u);
}

TEST_F(ComPtrTests, Test_com_ptr_t_move) {
    IMyLoggerPtr src = new ComObjMock;

    IMyLoggerPtr dst;
    dst = std::move(src);
    EXPECT_TRUE(dst);
    EXPECT_FALSE(src);

    EXPECT_EQ(GetRefCount(*dst), 1u);
}

TEST_F(ComPtrTests, Test_CComPtr_move) {
    CComPtr<IMyLogger> src(new ComObjMock);

    CComPtr<IMyLogger> dst;
    dst = std::move(src);
    EXPECT_TRUE(dst);
    EXPECT_FALSE(src);

    EXPECT_EQ(GetRefCount(*dst), 1u);
}

TEST_F(ComPtrTests, Test_com_ptr_t_AttachDetach) {
    IUnknownPtr ptr1(new ComObjMock);
    EXPECT_EQ(GetRefCount(*ptr1), 1u);

    IUnknownPtr ptr2;
    ptr2.Attach(ptr1);
    EXPECT_EQ(GetRefCount(*ptr1), 1u); // ref-count NOT increased
    EXPECT_EQ(GetRefCount(*ptr2), 1u); // ref-count NOT increased

    ptr1.Detach(); // prevent double-delete
}

TEST_F(ComPtrTests, Test_com_ptr_t_Release) {
    {
        // release with Release
        IUnknownPtr ptr1(new ComObjMock);
        EXPECT_EQ(GetRefCount(*ptr1), 1u);
        EXPECT_TRUE(ptr1);
        ptr1.Release();
        EXPECT_FALSE(ptr1);
    }
    {
        // release with nullptr assignment
        IUnknownPtr ptr2(new ComObjMock);
        EXPECT_EQ(GetRefCount(*ptr2), 1u);
        EXPECT_TRUE(ptr2);
        ptr2 = nullptr;
        EXPECT_FALSE(ptr2);
    }
}

TEST_F(ComPtrTests, Test_com_ptr_t_call) {
    IMyLoggerPtr ptr(new ComObjMock);
    EXPECT_TRUE(ptr);

    CHECK(ptr->Log(CComBSTR(L"Message")));
}

TEST_F(ComPtrTests, Test_com_ptr_t_cast_to_if) {
    // start with IUnknown
    IUnknownPtr ptr1(new ComObjMock);
    EXPECT_TRUE(ptr1);

    {
        // cast through ctor
        IMyLoggerPtr ptr2(ptr1);
        EXPECT_TRUE(ptr2);
    }
    {
        // cast through assignment
        IMyLoggerPtr ptr3;
        ptr3 = ptr1;
        EXPECT_TRUE(ptr3);
    }
}

TEST_F(ComPtrTests, Test_com_ptr_t_cast_from_if) {
    // start with custom interface
    IMyLoggerPtr ptr1(new ComObjMock);
    EXPECT_TRUE(ptr1);

    {
        // cast through ctor
        IUnknownPtr ptr2(ptr1);
        EXPECT_TRUE(ptr2);
    }
    {
        // cast through assignment
        IUnknownPtr ptr3;
        ptr3 = ptr1;
        EXPECT_TRUE(ptr3);
    }
}

TEST_F(ComPtrTests, Test_com_ptr_t_cast_CComPtr) {
    // start with _com_ptr_t
    IMyLoggerPtr ptr1(new ComObjMock);
    EXPECT_TRUE(ptr1);

    // cast to CComPtr<T>
    CComPtr<IMyLogger> ptr2(&*ptr1);
    EXPECT_TRUE(ptr2);

    // verify that ref-count become 2 after assigning to ptr2
    EXPECT_EQ(GetRefCount(*ptr1), 2u);
}

TEST_F(ComPtrTests, Test_com_ptr_t_operator_bool) {
    IMyLoggerPtr ptr1; // invalid nullptr object
    bool isValid1 = ptr1;
    EXPECT_FALSE(isValid1);
    EXPECT_TRUE(ptr1 == nullptr);
    EXPECT_FALSE(ptr1 != nullptr);

    IMyLoggerPtr ptr2(new ComObjMock); // valid object
    bool isValid2 = ptr2;
    EXPECT_TRUE(isValid2);
    EXPECT_FALSE(ptr2 == nullptr);
    EXPECT_TRUE(ptr2 != nullptr);
}

TEST_F(ComPtrTests, Test_com_ptr_t_null_ctor) {
    IUnknownPtr obj(nullptr);
}

TEST_F(ComPtrTests, Test_CComPtr_cast_com_ptr_t) {
    // start with CComPtr<T>
    CComPtr<IMyLogger> ptr1(new ComObjMock);
    EXPECT_TRUE(ptr1);

    // cast to _com_ptr_t
    IMyLoggerPtr ptr2(ptr1.p);
    EXPECT_TRUE(ptr2);

    // verify that ref-count become 2 after assigning to ptr2
    EXPECT_EQ(GetRefCount(*ptr1), 2u);
}

TEST_F(ComPtrTests, Test_com_ptr_t_Comparison) {
    // underlying test object
    IMyLoggerPtr obj1(new ComObjMock); // derived smart-ptr
    IUnknownPtr obj2 = obj1; // base smart-ptr
    IMyLogger* ptr1 = obj1; // derived ptr
    IUnknown* ptr2 = obj2; // base ptr

    // compare all combinations of _com_ptr_t against _com_ptr_t
    EXPECT_TRUE(obj1 == obj1);
    EXPECT_TRUE(obj1 == obj2);
    EXPECT_TRUE(obj2 == obj2);
    EXPECT_TRUE(obj2 == obj1);

    // compare all combinations of _com_ptr_t against raw pointer
    EXPECT_TRUE(obj1 == ptr1);
    EXPECT_TRUE(obj1 == ptr2);
    EXPECT_TRUE(obj2 == ptr2);
    EXPECT_TRUE(obj2 == ptr1);

    // compare all combinations of raw pointer against _com_ptr_t
    EXPECT_TRUE(ptr1 == obj1);
    EXPECT_TRUE(ptr1 == obj2);
    EXPECT_TRUE(ptr2 == obj2);
    EXPECT_TRUE(ptr2 == obj1);

    // no need for testing raw pointer against raw pointer comparison
}
