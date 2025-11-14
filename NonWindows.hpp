#pragma once
/* Minimal subset of COM defines for compatibility with non-Windows platforms.
   Does deliberately NOT expose any C++ standard library types through its API
   to decouple C++ standard library usage in client vs. server. */

#ifdef _WIN32
#error Header not intended for Windows platform
#endif

#define DECLSPEC_UUID(arg) 

#include <cassert>
#include <atomic>
#include <string>
#include <string.h> // for wcsdup
#include <codecvt>
#include <locale>
#include <iostream>
#include <type_traits>


/** Taken from guiddef.h. */
struct GUID {
    unsigned int   Data1;
    unsigned short Data2;
    unsigned short Data3;
    unsigned char  Data4[ 8 ];

    bool operator == (const GUID & other) const {
        return memcmp(this, &other, sizeof(GUID)) == 0;
    }
    bool operator < (const GUID & other) const {
        int diff = memcmp(this, &other, sizeof(GUID));
        return (diff < 0);
    }
};
static_assert(sizeof(GUID) == 16, "GUID not packed");

// __uuidof emulation
template<typename Q>
static GUID hold_uuidof () { return {}; }
#define DEFINE_UUIDOF_ID(Q, IID) template<> inline GUID hold_uuidof<Q>() { return IID; }
#define DEFINE_UUIDOF(Q) DEFINE_UUIDOF_ID(Q, IID_##Q)
#define __uuidof(Q) hold_uuidof<Q>()

typedef GUID           IID;
typedef unsigned int   DWORD;   ///< 32bit unsigned
typedef long           BOOL;
typedef unsigned char  BYTE;
typedef unsigned short USHORT;  ///< 16bit unsigned
typedef unsigned int   UINT;    ///< 32bit int
typedef unsigned int   ULONG;   ///< 32bit unsigned (cannot use 'long' since it's 64bit on 64bit Linux)
typedef int            LONG;    ///< 32bit int (cannot use 'long' since it can be 64bit)
typedef short  VARIANT_BOOL;    ///< boolean type that's natively marshaled to C# and Python
typedef wchar_t*       BSTR;    ///< zero terminated double-byte text string
typedef int32_t        HRESULT; ///< 32bit signed int (negative values indicate failure)
typedef void*          HWND;    ///< window handle
typedef unsigned long long  ULONGLONG; ///< 64bit unsigned
static_assert(sizeof(int) == 4, "int size not 32bit");
#define __int64       long long ///< 64bit int (cannot use typedef due to "unsigned __int64" code)

// Common HRESULT codes
// REF: https://learn.microsoft.com/en-us/windows/win32/seccrypto/common-hresult-values
#define S_OK           static_cast<int32_t>(0L)
#define S_FALSE        static_cast<int32_t>(1L)
#define E_BOUNDS       static_cast<int32_t>(0x8000000BL)
#define E_NOTIMPL      static_cast<int32_t>(0x80004001L)
#define E_NOINTERFACE  static_cast<int32_t>(0x80004002L)
#define E_POINTER      static_cast<int32_t>(0x80004003L)
#define E_ABORT        static_cast<int32_t>(0x80004004L)
#define E_FAIL         static_cast<int32_t>(0x80004005L)
#define E_UNEXPECTED   static_cast<int32_t>(0x8000FFFFL)
#define E_ACCESSDENIED static_cast<int32_t>(0x80070005L)
#define E_HANDLE       static_cast<int32_t>(0x80070006L)
#define E_OUTOFMEMORY  static_cast<int32_t>(0x8007000EL)
#define E_INVALIDARG   static_cast<int32_t>(0x80070057L)
#define E_NOT_SET      static_cast<int32_t>(0x80070490L)


enum CLSCTX { 
  CLSCTX_INPROC_SERVER           = 0x1,
  CLSCTX_INPROC_HANDLER          = 0x2,
  CLSCTX_LOCAL_SERVER            = 0x4,
  CLSCTX_REMOTE_SERVER           = 0x10,
  CLSCTX_NO_CODE_DOWNLOAD        = 0x400,
  CLSCTX_NO_CUSTOM_MARSHAL       = 0x1000,
  CLSCTX_ENABLE_CODE_DOWNLOAD    = 0x2000,
  CLSCTX_NO_FAILURE_LOG          = 0x4000,
  CLSCTX_DISABLE_AAA             = 0x8000,
  CLSCTX_ENABLE_AAA              = 0x10000,
  CLSCTX_FROM_DEFAULT_CONTEXT    = 0x20000,
  CLSCTX_ACTIVATE_32_BIT_SERVER  = 0x40000,
  CLSCTX_ACTIVATE_64_BIT_SERVER  = 0x80000,
  CLSCTX_ENABLE_CLOAKING         = 0x100000,
  CLSCTX_APPCONTAINER            = 0x400000,
  CLSCTX_ACTIVATE_AAA_AS_IU      = 0x800000,
  CLSCTX_PS_DLL                  = 0x80000000
};
#define CLSCTX_ALL              (CLSCTX_INPROC_SERVER| \
                                 CLSCTX_INPROC_HANDLER| \
                                 CLSCTX_LOCAL_SERVER| \
                                 CLSCTX_REMOTE_SERVER)

#define SUCCEEDED(hr) (((HRESULT)(hr)) >= 0)
#define FAILED(hr)    (((HRESULT)(hr)) < 0)

inline const char* InternalHresultToString(HRESULT hr) {
    switch (hr) {
    case S_OK:          return "S_OK";
    case S_FALSE:       return "S_FALSE";
    case E_BOUNDS:      return "E_BOUNDS";
    case E_NOTIMPL:     return "E_NOTIMPL";
    case E_NOINTERFACE: return "E_NOINTERFACE";
    case E_POINTER:     return "E_POINTER";
    case E_ABORT:       return "E_ABORT";
    case E_FAIL:        return "E_FAIL";
    case E_UNEXPECTED:  return "E_UNEXPECTED";
    case E_ACCESSDENIED:return "E_ACCESSDENIED";
    case E_HANDLE:      return "E_HANDLE";
    case E_OUTOFMEMORY: return "E_OUTOFMEMORY";
    case E_INVALIDARG:  return "E_INVALIDARG";
    case E_NOT_SET:     return "E_NOT_SET";
    default:            return "HRESULT error";
    }
}


inline void CHECK (HRESULT hr) {
    if (hr >= 0)
        return; // success

    const char* str = InternalHresultToString(hr);
    throw std::runtime_error(str);
}


/** API-compatible subset of the Microsoft _com_error class documented on https://docs.microsoft.com/en-us/cpp/cpp/com-error-class */
class _com_error {
public:
    _com_error(HRESULT hr) : m_hr(hr) {
        const char* str = InternalHresultToString(hr);

        size_t len = strlen(str);
        m_buffer.resize(len, L'\0');
        mbstowcs(const_cast<wchar_t*>(m_buffer.data()), str, len);
    }

    HRESULT Error() const noexcept {
        return m_hr;
    }

    const wchar_t* ErrorMessage() const noexcept {
        return m_buffer.c_str();
    }

private:
    HRESULT      m_hr = E_FAIL;
    std::wstring m_buffer;
};


/** API-compatible subset of the Microsoft _bstr_t class documented on https://docs.microsoft.com/en-us/cpp/cpp/bstr-t-class */
class _bstr_t {
public:
    _bstr_t() noexcept = default;
    
    _bstr_t(const _bstr_t& s) noexcept {
        Assign(s.m_str);
    }
    _bstr_t(const wchar_t* s) {
        Assign(s);
    }
    _bstr_t(wchar_t* s, bool copy) {
        if (copy)
            Assign(s);
        else
            m_str = s; // attach to string
    }

    ~_bstr_t() noexcept {
        Clear();
    }

    _bstr_t& operator=(const _bstr_t& s) noexcept {
        Assign(s.m_str);
        return *this;
    }
    _bstr_t& operator=(const wchar_t* s) {
        Assign(s);
        return *this;
    }
    
    _bstr_t& operator+=(const _bstr_t& s) {
        _bstr_t temp = *this + s;
        Assign(temp.m_str);
        return *this;
    }
    
    _bstr_t operator+(const _bstr_t& s) const {
        auto tmp = std::wstring(m_str) + std::wstring(s.m_str);
        _bstr_t result(tmp.c_str());
        return result;
    }

    operator wchar_t*() const noexcept {
        return m_str;
    }
    
    bool operator == (const _bstr_t& other) const noexcept {
        if (m_str && other.m_str)
            return wcscmp(m_str, other.m_str) == 0;

        return false;
    }
    bool operator != (const _bstr_t& other) const noexcept {
        return !operator == (other);
    }

    /** Returns string length excluding null termination. */
    unsigned int length() const noexcept {
        if (!m_str)
            return 0;

        return static_cast<unsigned int>(wcslen(m_str));
    }

    void Assign(const wchar_t* s) {
        if (s == m_str)
            return; // self-assignment

        Clear();
        
        if (s)
            m_str = wcsdup(s);
    }

    BSTR* GetAddress() {
        Clear();
        return &m_str;
    }

    void Attach(wchar_t* s) {
        Assign(s);
    }

    wchar_t* Detach() {
        wchar_t* tmp = m_str;
        m_str = nullptr;
        return tmp;
    }

private:
    void Clear() {
        if (!m_str)
            return;

        free(m_str);
        m_str = nullptr;
    }
    
    wchar_t* m_str = nullptr;
};
static_assert(sizeof(_bstr_t) == sizeof(wchar_t*), "_bstr_t size mismatch");


namespace ATL {

class CComBSTR {
public:
    CComBSTR () {
    }
    CComBSTR (const wchar_t* str) {
        if (str)
            m_str = wcsdup(str);
    }
    CComBSTR (int /*size*/, const wchar_t* str) {
        m_str = wcsdup(str);
    }
    CComBSTR (const CComBSTR & other) {
        m_str = other.Copy();
    }

    ~CComBSTR() {
        Empty();
    }

    void operator = (const CComBSTR & other) {
        if (&other == this)
            return;

        Empty();
        m_str = other.Copy();
    }

    /** Returns string length excluding null termination. */
    unsigned int Length () const {
        if (!m_str)
            return 0;

        return static_cast<unsigned int>(wcslen(m_str));
    }

    operator wchar_t* () const {
        return m_str;
    }

    wchar_t** operator & () {
        return &m_str;
    }
    
    void Attach(wchar_t* s) noexcept {
        if (s == m_str)
            return;

        Empty();
        m_str = s;
    }
    
    wchar_t* Detach () {
        wchar_t* tmp = m_str;
        m_str = nullptr;
        return tmp;
    }

    wchar_t* Copy () const {
        if (!m_str)
            return nullptr;

        return wcsdup(m_str);
    }
    
    void Empty() {
        if (!m_str)
            return;

        free(m_str);
        m_str = nullptr;
    }
    
    CComBSTR& operator+= (const wchar_t* other) {
        auto tmp = std::wstring(m_str) + std::wstring(other);
        operator=(tmp.c_str());
        return *this;
    }

    bool operator == (const CComBSTR& other) const {
        if (m_str && other.m_str)
            return wcscmp(m_str, other.m_str) == 0;

        return false;
    }
    bool operator != (const wchar_t* other) const {
        return !operator == (other);
    }

    wchar_t* m_str = nullptr;
};
static_assert(sizeof(CComBSTR) == sizeof(wchar_t*), "CComBSTR size mismatch");

template<typename T>
class CComPtr;
} // namespace ATL

template<typename T>
class _com_ptr_t;

extern "C" {
// interface ID values for well-known interfaces
static constexpr GUID IID_IUnknown       = {0x00000000,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static constexpr GUID IID_IMessageFilter = {0x00000016,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};

/** IUnknown base-class for Non-Windows platforms. */
struct IUnknown {
    IUnknown () {
    }

    virtual ~IUnknown() {
    }

    /** Cast method. */
    virtual HRESULT QueryInterface (const GUID & iid, /*[out]*/void **obj) = 0;

    virtual ULONG AddRef () = 0;

    virtual ULONG Release () = 0;
};
} // extern "C"
DEFINE_UUIDOF(IUnknown)


// error handler required by generated wrapper API headers
inline void _com_issue_errorex(HRESULT hr, IUnknown*, const IID &) {
    throw _com_error(hr);
}

template <class BASE>
class CComObject : public BASE {
public:
    static HRESULT CreateInstance (CComObject<BASE> ** arg) {
        assert(arg);
        assert(!*arg);

        auto* ptr = new CComObject<BASE>();
        HRESULT hr = ptr->FinalConstruct();
        if (FAILED(hr)) {
            delete ptr;
            ptr = nullptr;
        }

        *arg = ptr;
        return hr;
    }
};

template <class BASE>
class CComContainedObject : public BASE {
public:
    CComContainedObject (IUnknown* pOuterUnknown) : m_pOuterUnknown(pOuterUnknown) {
    }
    
    // forward reference-counting & QI to controlling outer
    ULONG AddRef () override {
        return m_pOuterUnknown->AddRef();
    }
    ULONG Release () override {
        return m_pOuterUnknown->Release();
    }
    HRESULT QueryInterface (const GUID & iid, /*out*/void **obj) override {
        return m_pOuterUnknown->QueryInterface(iid, obj);
    }
    HRESULT InternalQueryInterface (const GUID & iid, /*out*/void **obj) {
        return BASE::QueryInterface(iid, obj);
    }
    
private:
    IUnknown* m_pOuterUnknown = nullptr;
};

template <class BASE>
class CComAggObject : public IUnknown {
public:
    CComAggObject (IUnknown* pOuterUnknown) : m_contained(pOuterUnknown) {
    }
    ~CComAggObject () {
    }
    
    ULONG AddRef () override {
        return ++m_ref;
    }
    ULONG Release () override {
        ULONG ref = --m_ref;
        if (!ref)
            delete this;
        
        return ref;
    }
    HRESULT QueryInterface (const GUID & iid, /*out*/void **obj) override {
        if (iid == __uuidof(IUnknown)) {
            // special handling of IUnknown
            *obj = static_cast<IUnknown*>(this);
            AddRef();
            return S_OK;
        } else {
            return m_contained.InternalQueryInterface(iid, obj);
        }
    }

    static HRESULT CreateInstance (IUnknown* unkOuter, CComAggObject<BASE> ** arg) {
        assert(unkOuter);
        assert(arg);
        assert(!*arg);

        auto* ptr = new CComAggObject<BASE>(unkOuter);
        HRESULT hr = ptr->m_contained.FinalConstruct();
        if (FAILED(hr)) {
            delete ptr;
            ptr = nullptr;
        }

        *arg = ptr;
        return hr;
    }
    
    CComContainedObject<BASE> m_contained;
private:
    std::atomic<ULONG>        m_ref {0};
};


/** std::vector alternative to avoid C++ standard library dependency. */
template <class T>
class Buffer {
public:
    Buffer(size_t size = 0) : m_size(size) {
        if (size > 0) {
            m_ptr = Allocate(size);
            m_owning = true;
        }
    }
    Buffer(const Buffer& other, bool deep_copy) : m_size(other.m_size) {
        if (deep_copy) {
            m_ptr = Allocate(m_size);
            m_owning = true;

            for (size_t i = 0; i < m_size; i++)
                m_ptr[i] = other.m_ptr[i];
        } else {
            m_ptr = other.m_ptr;
            m_owning = false;
        }
    }

    ~Buffer() {
        if (m_ptr && m_owning)
            Free(m_ptr, m_size);
        m_ptr = nullptr;
    }

    size_t size() const {
        return m_size;
    }

    T * data() {
        return m_ptr;
    }

    T& operator [](size_t idx) {
        return m_ptr[idx];
    }
    const T& operator [](size_t idx) const {
        return m_ptr[idx];
    }

    void resize (size_t size, T val = T()) noexcept {
        assert(m_owning);

        // allocate new buffer
        T * new_ptr = nullptr;
        if (size > 0) {
            new_ptr = Allocate(size);
            m_owning = true;

            for (size_t i = 0; i < std::min(m_size, size); ++i)
                new_ptr[i] = m_ptr[i];

            for (size_t i = std::min(m_size, size); i < size; ++i)
                new_ptr[i] = val;
        }
        // delete old buffer
        if (m_ptr && m_owning)
            Free(m_ptr, m_size);
        // commit changes
        m_ptr = new_ptr;
        m_size = size;
    }

    Buffer(const Buffer& other) = delete;
    Buffer(Buffer&&) = delete;
    Buffer& operator = (const Buffer&) = delete;
    Buffer& operator = (Buffer&&) = delete;

private:
    static T* Allocate(size_t size) {
        auto* ptr = (T*)malloc(sizeof(T)*size);
        for (size_t i = 0; i < size; i++)
            new (&ptr[i]) T();

        return ptr;
    }
    
    static void Free (T* ptr, size_t size) {
        for (size_t i = 0; i < size; i++)
            ptr[i].~T();
        free(ptr);
    }
    
    size_t m_size = 0;
    T  *   m_ptr = nullptr;
    bool   m_owning = true;
};


/** Internal class that SHALL ONLY be accessed through _com_ptr_t<T> or CComPtr<T> to preserve Windows compatibility. */
class IUnknownFactory {
    template<typename T>
    friend class _com_ptr_t;
    template<typename T>
    friend class ATL::CComPtr;

private:
    typedef HRESULT(*Factory)(IUnknown*, IUnknown**);
    
    struct Entry {
        GUID          clsid{};
        ATL::CComBSTR name;
        Factory       factory = nullptr;
    };

    /** Create COM class based on "[<Program>.]<Component>[.<Version>]" ProgID string. */
    static IUnknown* CreateInstance (std::wstring class_name, IUnknown* outer) {
        // remove "<Program>." prefix and ".<Version>" suffix if present
        size_t idx1 = class_name.find(L'.');
        if (idx1 != std::wstring::npos) {
            std::wstring suffix = class_name.substr(idx1+1); // "<Component>.<Version>" or "<Version>"
            size_t idx2 = suffix.find(L'.');

            if (idx2 != std::wstring::npos) {
                // input contain two '.'s, keep center part
                class_name = suffix.substr(0,idx2); // "<Component>"
            } else {
                // input contain one '.'. check if suffix is a number
                auto version = wcstol(suffix.c_str(), nullptr, /*base*/10);
                if (version != 0) {
                    class_name = class_name.substr(0, idx1);
                } else {
                    class_name = suffix;
                }
            }
        }

        for (size_t i = 0; i < Factories().size(); i++) {
            const auto & elm = Factories()[i];
            if (elm.name == ATL::CComBSTR(class_name.c_str())) {
                IUnknown* obj = nullptr;
                HRESULT hr = elm.factory(outer, &obj);
                assert(hr == S_OK);
                if (SUCCEEDED(hr))
                    return obj;
            }
        }

        std::wcerr << L"CoCreateInstance error: Unknown class " << class_name << std::endl;
        assert(false);
        return nullptr;
    }

    /** Create COM class based on CLSID. */
    static IUnknown* CreateInstance (GUID clsid, IUnknown* outer) {
        for (size_t i = 0; i < Factories().size(); i++) {
            const auto & elm = Factories()[i];
            if (elm.clsid == clsid) {
                IUnknown* obj = nullptr;
                HRESULT hr = elm.factory(outer, &obj);
                assert(hr == S_OK);
                if (SUCCEEDED(hr))
                    return obj;
            }
        }

        char guid_str[39] = {};
        snprintf(guid_str, sizeof(guid_str), "{%08x-%04x-%04x-%02x%02x-%02x%02x%02x%02x%02x%02x}",
            clsid.Data1, clsid.Data2, clsid.Data3,
            clsid.Data4[0], clsid.Data4[1], clsid.Data4[2], clsid.Data4[3],
            clsid.Data4[4], clsid.Data4[5], clsid.Data4[6], clsid.Data4[7]);

        std::cerr << "CoCreateInstance error: Unknown clsid " << guid_str << std::endl;
        assert(false);
        return nullptr;
    }
    
public:
    template <class CLS>
    static const char* RegisterClass(GUID clsid, const char * class_name) {

        std::wstring w_class_name(strlen(class_name), L'\0');
        mbstowcs(const_cast<wchar_t*>(w_class_name.data()), class_name, w_class_name.size());

        //printf("IUnknownFactory::RegisterClass(%s)\n", class_name);
        size_t prev_size = Factories().size();
        Factories().resize(prev_size + 1, {clsid, w_class_name.c_str(), CreateClass<CLS>});
        return class_name; // pass-through name
    }

private:
    template <class CLS>
    static HRESULT CreateClass (IUnknown* outer, IUnknown** obj) {
        if (outer) {
            // create an object (with ref. count zero)
            CComAggObject<CLS> * tmp = nullptr;
            HRESULT hr = CComAggObject<CLS>::CreateInstance(outer, &tmp);
            if (FAILED(hr))
                return hr;

            tmp->AddRef(); // incr. ref-count to one
            *obj = tmp;
            return hr;
        } else {
            // create an object (with ref. count zero)
            CComObject<CLS> * tmp = nullptr;
            HRESULT hr = CComObject<CLS>::CreateInstance(&tmp);
            if (FAILED(hr))
                return hr;

            tmp->AddRef(); // incr. ref-count to one
            *obj = tmp;
            return hr;
        }
    }

    static Buffer<Entry> & Factories ();
};

#define OBJECT_ENTRY_AUTO(clsid, cls) \
    __attribute__((weak)) const char* tmp_factory_##cls = IUnknownFactory::RegisterClass<cls>(clsid, #cls);


/** Mostly API-compatible subset of the Microsoft _com_ptr_t class documented on https://docs.microsoft.com/en-us/cpp/cpp/com-ptr-t-class */
template <class T>
class _com_ptr_t {
public:
    _com_ptr_t() noexcept = default;

    _com_ptr_t(decltype(nullptr)) noexcept {
    }

    /** Matching COM ptr ctor. */
    _com_ptr_t (T * ptr, bool addref = true) : m_ptr(ptr) {
        assert(m_ptr && "_com_ptr_t::ctor nullptr.");
        if (addref)
            m_ptr->AddRef();
    }
    /** Matching smart-ptr ctor. */
    _com_ptr_t (const _com_ptr_t & other) : m_ptr(other.m_ptr) {
        if (m_ptr)
            m_ptr->AddRef();
    }
    
    /** Casting smart-ptr ctor. */
    template<typename Q, std::enable_if_t<!std::is_same<Q, T>::value, bool> = true>   // call _com_ptr_t ctor instead
    _com_ptr_t(const _com_ptr_t<Q>& ptr) {
        assert(ptr && "_com_ptr_t::ctor nullptr.");
        HRESULT hr = ptr.QueryInterface(__uuidof(T), &m_ptr);
        assert(((hr == S_OK) || (hr == E_NOINTERFACE)) && "_com_ptr_t::ctor cast failure.");
        (void)hr; // mute unreferenced variable warning
    }
    /** Casting COM ptr ctor. */
    template<typename Q, std::enable_if_t<!(
           std::is_same<Q, T>::value          // call T* ctor instead
        || std::is_same<Q, _com_ptr_t>::value // call _com_ptr_t ctor instead
        ), bool> = true>
    _com_ptr_t(Q * ptr) {
        assert(ptr && "_com_ptr_t::ctor nullptr.");
        HRESULT hr = ptr->QueryInterface(__uuidof(T), reinterpret_cast<void**>(&m_ptr));
        assert(((hr == S_OK) || (hr == E_NOINTERFACE)) && "_com_ptr_t::ctor cast failure.");
        (void)hr; // mute unreferenced variable warning
    }

    ~_com_ptr_t () {
        if (m_ptr)
            Release();
    }

    /** Nullptr assignment. */
    void operator = (std::nullptr_t) {
        if (m_ptr)
            Release();
    }
    /** Smart-ptr assignment. */
    void operator = (const _com_ptr_t & other) {
        if (m_ptr != other.m_ptr)
            _com_ptr_t(other).Swap(*this);
    }

    _com_ptr_t& operator=(_com_ptr_t&& other) noexcept {
        if (m_ptr != other.m_ptr) {
            if (m_ptr)
                Release();
            std::swap(m_ptr, other.m_ptr);
        }

        return *this;
    }

    template <class Q>
    HRESULT QueryInterface (const IID& iid, Q** arg) const {
        if (!m_ptr)
            return E_POINTER;

        return m_ptr->QueryInterface(iid, reinterpret_cast<void**>(arg));
    }

    /** Take over ownership (does not incr. ref-count). */
    void Attach (T * ptr) {
        if (m_ptr)
            Release();
        m_ptr = ptr;
        // no AddRef
    }
    /** Release ownership (does not decr. ref-count). */
    T* Detach () {
        T * tmp = m_ptr;
        m_ptr = nullptr;
        // no Release
        return tmp;
    }

    void Release () {
        assert(m_ptr && "_com_ptr_t::Release nullptr.");

        m_ptr->Release();
        m_ptr = nullptr;
    }

    operator T*() const noexcept {
        return m_ptr;
    }

    operator T& () const {
        assert(m_ptr && "_com_ptr_t::operator& nullptr.");
        return *m_ptr;
    }

    T& operator* () const {
        assert(m_ptr && "_com_ptr_t::operator* nullptr.");
        return *m_ptr;
    }        

    T** operator& () noexcept {
        if (m_ptr)
            Release();
        return &m_ptr;
    }

    T* operator-> () const {
        assert(m_ptr && "_com_ptr_t::operator-> nullptr.");
        return m_ptr;
    }

    bool operator==(T* p) const {
        if (m_ptr == p)
            return true;
        return CompareUnknown(p) == 0;
    }

    HRESULT CreateInstance (const GUID& clsid, IUnknown* outer = nullptr, DWORD context = CLSCTX_ALL) noexcept {
        (void)context;

        _com_ptr_t<IUnknown> tmp1(IUnknownFactory::CreateInstance(clsid, outer), /*addref*/false);
        if (!tmp1)
            return E_FAIL;

        _com_ptr_t tmp2 = tmp1; // cast
        if (!tmp2)
            return E_NOINTERFACE;

        Swap(tmp2);
        return S_OK;
    }

    HRESULT CreateInstance(const wchar_t* name, IUnknown* outer = nullptr, DWORD context = CLSCTX_ALL) noexcept {
        (void)context;

        if (!name)
            return E_INVALIDARG;

        _com_ptr_t<IUnknown> tmp1(IUnknownFactory::CreateInstance(name, outer), /*addref*/false);
        if (!tmp1)
            return E_FAIL;

        _com_ptr_t tmp2 = tmp1; // cast
        if (!tmp2)
            return E_NOINTERFACE;

        Swap(tmp2);
        return S_OK;
    }

private:
    T * m_ptr = nullptr;

    void Swap (_com_ptr_t & other) {
        T* tmp = m_ptr;
        m_ptr = other.m_ptr;
        other.m_ptr = tmp;
    }

    ptrdiff_t CompareUnknown(T * p) const {
        static_assert(std::is_base_of<IUnknown, T>::value, "_com_ptr_t::CompareUnknown: T must inherit from IUnknown");

        IUnknown* pu1 = nullptr;
        IUnknown* pu2 = nullptr;

        if (m_ptr) {
            HRESULT hr = m_ptr->QueryInterface(__uuidof(IUnknown), reinterpret_cast<void**>(&pu1));
            (void)hr;
            assert(SUCCEEDED(hr) && "_com_ptr_t::CompareUnknown cast failed");
            pu1->Release();
        }

        if (p) {
            HRESULT hr = p->QueryInterface(__uuidof(IUnknown), reinterpret_cast<void**>(&pu2));
            (void)hr;
            assert(SUCCEEDED(hr) && "_com_ptr_t::CompareUnknown cast failed");
            pu2->Release();
        }

        return (pu1 - pu2);
    }
};

// Support _COM_SMARTPTR_TYPEDEF defines in generated wrapper API headers
#define _COM_SMARTPTR_TYPEDEF(Interface, IID) typedef _com_ptr_t<Interface> Interface ## Ptr

// IUnknown smart-ptr define
_COM_SMARTPTR_TYPEDEF(IUnknown, __uuidof(IUnknown));


namespace ATL {

/** COM smart-pointer object. */
template <class T>
class CComPtr {
public:
    CComPtr (T * ptr = nullptr) : p(ptr) {
        if (p)
            p->AddRef();
    }
    CComPtr (const CComPtr & other) : p(other.p) {
        if (p)
            p->AddRef();
    }

    ~CComPtr () {
        if (p)
            p->Release();
        p = nullptr;
    }

    void operator = (T * other) {
        if (p != other)
            CComPtr(other).Swap(*this);
    }
    void operator = (const CComPtr & other) {
        if (p != other.p)
            CComPtr(other).Swap(*this);
    }
    template <typename U>
    void operator = (const CComPtr<U> & other) {
        CComPtr tmp;
        other.QueryInterface(&tmp);
        Swap(tmp);
    }

    CComPtr& operator=(CComPtr&& other) noexcept {
        if (p != other.p) {
            if (p)
                Release();
            std::swap(p, other.p);
        }

        return *this;
    }

    template <class Q>
    HRESULT QueryInterface (Q** arg) const {
        if (!p)
            return E_POINTER;
        if (!arg)
            return E_POINTER;
        if (*arg)
            return E_POINTER;
        
        return p->QueryInterface(__uuidof(Q), reinterpret_cast<void**>(arg));
    }

    bool IsEqualObject (IUnknown * other) const {
        CComPtr<IUnknown> this_obj;
        this_obj = *this;
        return this_obj == other;
    }

    HRESULT CopyTo (T** arg) {
        if (!arg)
            return E_POINTER;
        if (*arg)
            return E_POINTER; // input must be pointer to nullptr

        *arg = p;
        if (p)
            p->AddRef();

        return S_OK;
    }

    HRESULT CoCreateInstance (std::wstring name, IUnknown* outer = NULL, DWORD context = CLSCTX_ALL) {
        (void)context;

        IUnknown* tmp0 = IUnknownFactory::CreateInstance(name, outer); // RefCount=1
        if (!tmp0)
            return E_FAIL;

        CComPtr<IUnknown> tmp1;
        tmp1.Attach(tmp0);

        CComPtr tmp2;
        tmp2 = tmp1; // cast
        if (!tmp2)
            return E_FAIL;

        Swap(tmp2);
        return S_OK;
    }

    HRESULT CoCreateInstance (GUID clsid, IUnknown* outer = NULL, DWORD context = CLSCTX_ALL) {
        (void)context;
        IUnknown* tmp0 = IUnknownFactory::CreateInstance(clsid, outer); // RefCount=1
        if (!tmp0)
            return E_FAIL;

        CComPtr<IUnknown> tmp1;
        tmp1.Attach(tmp0);

        CComPtr tmp2;
        tmp2 = tmp1; // cast
        if (!tmp2)
            return E_FAIL;

        Swap(tmp2);
        return S_OK;
    }

    /** Take over ownership (does not incr. ref-count). */
    void Attach (T * ptr) {
        if (p)
            p->Release();
        p = ptr;
        // no AddRef
    }
    /** Release ownership (does not decr. ref-count). */
    T* Detach () {
        T * tmp = p;
        p = nullptr;
        // no Release
        return tmp;
    }

    void Release () {
        if (!p)
            return;

        p->Release();
        p = nullptr;
    }

    operator T*() const {
        return p;
    }
    T* operator -> () const {
        assert(p && "CComPtr::operator -> nullptr.");
        return p;
    }
    T** operator & () {
        return &p;
    }

    T * p = nullptr;

protected:
    void Swap (CComPtr & other) {
        T* tmp = p;
        p = other.p;
        other.p = tmp;
    }    
};

template <class T>
using CComQIPtr = CComPtr<T>;

template <class T>
struct CComSafeArray;

} // namespace ATL

/** Internal class that SHALL ONLY be accessed through CComSafeArray<T> to preserve Windows compatibility. */
struct SAFEARRAY {
    template<typename T>
    friend struct ATL::CComSafeArray;

private:
    enum TYPE {
        TYPE_EMPTY,
        TYPE_DATA,
        TYPE_STRINGS,
        TYPE_POINTERS,
    };

    SAFEARRAY (TYPE t) : type(t), elm_size(sizeof(void*)) {
        assert(t == TYPE_STRINGS || t == TYPE_POINTERS);
    }
    SAFEARRAY (unsigned int _elm_size, unsigned int count) : type(TYPE_DATA), data(_elm_size*count), elm_size(_elm_size) {
    }
    SAFEARRAY(const SAFEARRAY& other, bool deep_copy) : type(other.type), data(other.data, deep_copy), strings(other.strings, deep_copy), pointers(other.pointers, deep_copy), elm_size(other.elm_size) {
    }

    ~SAFEARRAY() {
    }

    SAFEARRAY () = delete;
    SAFEARRAY& operator = (const SAFEARRAY&) = delete;
    
    static SAFEARRAY* Create(TYPE t) {
        auto* ptr = (SAFEARRAY*)malloc(sizeof(SAFEARRAY));
        new (ptr) SAFEARRAY(t);
        return ptr;
    }
    static SAFEARRAY* Create(unsigned int _elm_size, unsigned int count) {
        auto* ptr = (SAFEARRAY*)malloc(sizeof(SAFEARRAY));
        new (ptr) SAFEARRAY(_elm_size, count);
        return ptr;
    }
    static SAFEARRAY* Create(const SAFEARRAY& other, bool deep_copy = true) {
        auto* ptr = (SAFEARRAY*)malloc(sizeof(SAFEARRAY));
        new (ptr) SAFEARRAY(other, deep_copy);
        return ptr;
    }
    
    static void Destroy(SAFEARRAY* obj) {
        obj->~SAFEARRAY();
        free(obj);
    }

    const TYPE                     type = TYPE_EMPTY; ///< \todo: Replace with std::variant when upgrading to C++17
    Buffer<unsigned char>          data;
    Buffer<ATL::CComBSTR>          strings;
    Buffer<ATL::CComPtr<IUnknown>> pointers;
    const unsigned int             elm_size = 0;
};


namespace ATL {

template <typename T>
struct CComTypeWrapper {
    typedef T type; // default mapping
};
template <>
struct CComTypeWrapper<BSTR> {
    typedef CComBSTR type; // map BSTR/wchar_t* to CComBSTR
};
template <>
struct CComTypeWrapper<IUnknown*> {
    typedef CComPtr<IUnknown> type; // wrap IUnknown* in CComPtr
};

template <class T>
struct CComSafeArray {
    CComSafeArray () {
    }

    CComSafeArray (UINT size) {
        m_ptr = SAFEARRAY::Create(sizeof(T), size);
    }

    CComSafeArray (SAFEARRAY * obj) {
        if (obj) {
            assert(obj->elm_size == sizeof(T));
            m_ptr = SAFEARRAY::Create(*obj);
        }
    }

    ~CComSafeArray () {
        Destroy();
    }

    CComSafeArray (const CComSafeArray&) = delete;
    CComSafeArray& operator = (const CComSafeArray&) = delete;

    CComSafeArray (CComSafeArray&& other) {
        std::swap(m_ptr, other.m_ptr);
    }
    CComSafeArray& operator = (CComSafeArray&& other) {
        std::swap(m_ptr, other.m_ptr);
        return *this;
    }

    HRESULT Destroy() {
        if (m_ptr) {
            SAFEARRAY::Destroy(m_ptr);
            m_ptr = nullptr;
        }
        return S_OK;
    }

    HRESULT Attach (SAFEARRAY * obj) {
        assert(obj);
        assert(obj->elm_size == sizeof(T));
        Destroy();
        m_ptr = obj;
        return S_OK;
    }
    SAFEARRAY* Detach () {
        SAFEARRAY* tmp = m_ptr;
        m_ptr = nullptr;
        return tmp;
    }

    operator SAFEARRAY* () {
        return m_ptr;
    }

    typename CComTypeWrapper<T>::type& GetAt (int idx) const {
        assert(m_ptr);
        assert(m_ptr->type == SAFEARRAY::TYPE_DATA);
        unsigned char * ptr = &m_ptr->data[idx*m_ptr->elm_size];
        return reinterpret_cast<T&>(*ptr);
    }

    typename CComTypeWrapper<T>::type& operator [] (int idx) {
        return GetAt(idx);
    }

    HRESULT SetAt (int idx, const T& val, bool copy = true) {
        (void)copy; // mute unreferenced argument warning
        assert(m_ptr);
        assert(m_ptr->type == SAFEARRAY::TYPE_DATA);
        assert(sizeof(T) == m_ptr->elm_size);
        unsigned char * ptr = &m_ptr->data[idx*m_ptr->elm_size];
        reinterpret_cast<T&>(*ptr) = val;
        return S_OK;
    }

    HRESULT Add (const typename CComTypeWrapper<T>::type& t, BOOL copy = true) {
        (void)copy; // mute unreferenced argument warning
        
        if (!m_ptr)
            m_ptr = SAFEARRAY::Create(sizeof(T), 0); // lazy initialization

        assert(m_ptr->type == SAFEARRAY::TYPE_DATA);
        assert(sizeof(T) == m_ptr->elm_size);
        const size_t prev_size = m_ptr->data.size();
        m_ptr->data.resize(prev_size + sizeof(T), 0);
        reinterpret_cast<T&>(m_ptr->data[prev_size]) = t;
        return S_OK;
    }

    unsigned int GetCount () const {
        assert(m_ptr);
        assert(m_ptr->type == SAFEARRAY::TYPE_DATA);
        return static_cast<unsigned int>(m_ptr->data.size()/m_ptr->elm_size);
    }
    
    /** Internal function. Do NOT call unless you know what you're doing. */
    static typename CComTypeWrapper<T>::type* InternalDataPointer(SAFEARRAY* obj) {
        CComSafeArray<T> sa;
        sa.Attach(obj);
        typename CComTypeWrapper<T>::type* ptr = &sa.GetAt(0);
        sa.Detach();
        return ptr;
    }

    /** Internal function. Do NOT call unless you know what you're doing. */
    static unsigned int InternalElementCount(SAFEARRAY* obj) {
        CComSafeArray<T> sa;
        sa.Attach(obj);
        unsigned int count = sa.GetCount();
        sa.Detach();
        return count;
    }
    
    /** Internal function. Do NOT call unless you know what you're doing. */
    static SAFEARRAY* InternalShallowCopy(const SAFEARRAY& obj) {
        return SAFEARRAY::Create(obj, /*deep copy*/false);
    }

    SAFEARRAY* m_ptr = nullptr;
};
// Template specializations. Implemented in cpp file.
template <> CComSafeArray<BSTR>::CComSafeArray (UINT size);
template <> CComSafeArray<IUnknown*>::CComSafeArray (UINT size);
template <> CComTypeWrapper<BSTR>::type& CComSafeArray<BSTR>::GetAt (int idx) const;
template <> CComTypeWrapper<IUnknown*>::type& CComSafeArray<IUnknown*>::GetAt (int idx) const;
template <> HRESULT CComSafeArray<IUnknown*>::SetAt (int idx, IUnknown* const& val, bool copy);
template <> HRESULT                      CComSafeArray<BSTR>::Add (const typename CComTypeWrapper<BSTR>::type& t, BOOL copy);
template <> HRESULT                 CComSafeArray<IUnknown*>::Add (const typename CComTypeWrapper<IUnknown*>::type& t, BOOL copy);
template <> unsigned int CComSafeArray<BSTR>::GetCount () const;
template <> unsigned int CComSafeArray<IUnknown*>::GetCount () const;


// COM calling convention (use default on non-Windows)
#define STDMETHODCALLTYPE

#define ATL_NO_VTABLE 

// QueryInterface support macros
#define BEGIN_COM_MAP(CLASS)         HRESULT QueryInterface (const GUID & iid, /*out*/void **obj) override { \
                                           static_assert(std::is_same_v<CLASS, std::remove_pointer_t<decltype(this)>>, \
                                               "Argument to BEGIN_COM_MAP doesn't match name of surrounding class."); \
                                           *obj = nullptr;

#define COM_INTERFACE_ENTRY(INTERFACE) if (iid == __uuidof(INTERFACE)) \
                                           *obj = static_cast<INTERFACE*>(this); \
                                       else
#define COM_INTERFACE_ENTRY_AGGREGATE(INTERFACE, punk) \
                                       if (iid == __uuidof(INTERFACE)) \
                                           *obj = static_cast<INTERFACE*>(&punk->m_contained); \
                                       else
#define END_COM_MAP()                  if (iid == __uuidof(IUnknown)) \
                                           *obj = static_cast<IUnknown*>(this); \
                                       else \
                                           return E_NOINTERFACE; \
                                       AddRef(); \
                                       return S_OK; \
                                     } \
                                     ULONG AddRef () override { \
                                         assert((m_ref < 0xFFFF) && "IUnknown::AddRef negative ref count."); \
                                         return ++m_ref; \
                                     } \
                                     ULONG Release () override { \
                                         ULONG ref = --m_ref; \
                                         assert((m_ref < 0xFFFF) && "IUnknown::Release negative ref count."); \
                                         if (!ref) \
                                             delete this; \
                                         return ref; \
                                     } \
                                     std::atomic<ULONG> m_ref {0};

#define DECLARE_PROTECT_FINAL_CONSTRUCT()

#define DECLARE_REGISTRY_RESOURCEID(dummy)


class CComSingleThreadModel {};
class CComMultiThreadModel {};

template <class ThreadModel>
class CComObjectRootEx {
public:
    HRESULT FinalConstruct() {
        return S_OK;
    }
};

template <class T, const GUID* pclsid = nullptr>
class CComCoClass {
};


} // namespace ATL

#ifndef _ATL_NO_AUTOMATIC_NAMESPACE
  using namespace ATL;
#endif
