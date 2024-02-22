#pragma once
/* Minimal subset of COM defines for compatibility with non-Windows platforms. */

#ifdef _WIN32
#error Header not intended for Windows platform
#endif

#define DECLSPEC_UUID(arg) 

#include <memory>
#include <cassert>
#include <atomic>
#include <vector>
#include <string>
#include <string.h> // for wcsdup
#include <codecvt>
#include <locale>
#include <map>
#include <iostream>


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
static_assert(sizeof(int) == 4, "int size not 32bit");

// Common HRESULT codes
// REF: https://msdn.microsoft.com/en-us/library/windows/desktop/aa378137.aspx
#define S_OK           static_cast<int32_t>(0L)
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


inline const char* InternalHresultToString(HRESULT hr) {
    switch (hr) {
    case S_OK:          return "S_OK";
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

    bool operator == (const wchar_t* other) const {
        if (m_str && other)
            return wcscmp(m_str, other) == 0;

        return false;
    }
    bool operator != (const wchar_t* other) const {
        return !operator == (other);
    }

    wchar_t* m_str = nullptr;
};
static_assert(sizeof(CComBSTR) == sizeof(wchar_t*), "CComBSTR size mismatch");

} // namespace ATL

extern "C" {
// REF: https://msdn.microsoft.com/en-us/library/windows/desktop/ms682521.aspx
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


class IUnknownFactory {
public:
    typedef IUnknown* (*Factory)();

    /** Create COM class based on "[<Program>.]<Component>[.<Version>]" ProgID string. */
    static IUnknown* CreateInstance (std::string class_name) {
        // remove "<Program>." prefix and ".<Version>" suffix if present
        size_t idx1 = class_name.find('.');
        if (idx1 != std::string::npos) {
            std::string suffix = class_name.substr(idx1+1); // "<Component>.<Version>" or "<Version>"
            size_t idx2 = suffix.find('.');

            if (idx2 != std::string::npos) {
                // input contain two '.'s, keep center part
                class_name = suffix.substr(0,idx2); // "<Component>"
            } else {
                // input contain one '.'. check if suffix is a number
                auto version = strtol(suffix.c_str(), nullptr, /*base*/10);
                if (version != 0) {
                    class_name = class_name.substr(0, idx1);
                } else {
                    class_name = suffix;
                }
            }
        }

        for (const auto & elm : Factories()) {
            if (elm.first.second == class_name) {
                return elm.second();
            }
        }

        std::cerr << "CoCreateInstance error: Unknown class " << class_name << std::endl;
        assert(false);
        return nullptr;
    }

    /** Create COM class based on CLSID. */
    static IUnknown* CreateInstance (GUID clsid) {
        for (const auto & elm : Factories()) {
            if (elm.first.first == clsid) {
                return elm.second();
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
    
    template <class CLS>
    static const char* RegisterClass(GUID clsid, const char * class_name) {
        Factories()[{clsid,class_name}] = CreateClass<CLS>;
        return class_name; // pass-through name
    }

private:
    template <class CLS>
    static IUnknown* CreateClass () {
        return new CLS;
    }

    static std::map<std::pair<GUID,std::string>, Factory> & Factories ();
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
        ptr.QueryInterface(__uuidof(T), &m_ptr);
        assert(m_ptr && "_com_ptr_t::ctor cast failure.");
    }
    /** Casting COM ptr ctor. */
    template<typename Q, std::enable_if_t<!(
           std::is_same<Q, T>::value          // call T* ctor instead
        || std::is_same<Q, _com_ptr_t>::value // call _com_ptr_t ctor instead
        ), bool> = true>
    _com_ptr_t(Q * ptr) {
        assert(ptr && "_com_ptr_t::ctor nullptr.");
        ptr->QueryInterface(__uuidof(T), reinterpret_cast<void**>(&m_ptr));
        assert(m_ptr && "_com_ptr_t::ctor cast failure.");
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

    HRESULT CreateInstance (const GUID& clsid, IUnknown* outer = nullptr, DWORD context = CLSCTX_ALL) noexcept {
        (void)outer;
        (void)context;

        _com_ptr_t<IUnknown> tmp1(IUnknownFactory::CreateInstance(clsid));
        if (!tmp1)
            return E_FAIL;

        _com_ptr_t tmp2 = tmp1; // cast
        if (!tmp2)
            return E_NOINTERFACE;

        Swap(tmp2);
        return S_OK;
    }

    HRESULT CreateInstance(const wchar_t* name, IUnknown* outer = nullptr, DWORD context = CLSCTX_ALL) noexcept {
        (void)outer;
        (void)context;

        if (!name)
            return E_INVALIDARG;

        // convert name to ASCII
        std::string a_name(wcslen(name), '\0');
        wcstombs(const_cast<char*>(a_name.data()), name, a_name.size());

        _com_ptr_t<IUnknown> tmp1(IUnknownFactory::CreateInstance(a_name));
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
        (void)outer;
        (void)context;

        // convert name to ASCII
        std::string a_name(name.size(), '\0');
        wcstombs(const_cast<char*>(a_name.data()), name.c_str(), a_name.size());

        CComPtr<IUnknown> tmp1(IUnknownFactory::CreateInstance(a_name));
        if (!tmp1)
            return E_FAIL;

        CComPtr tmp2;
        tmp2 = tmp1; // cast
        if (!tmp2)
            return E_FAIL;

        Swap(tmp2);
        return S_OK;
    }

    HRESULT CoCreateInstance (GUID clsid, IUnknown* outer = NULL, DWORD context = CLSCTX_ALL) {
        (void)outer;
        (void)context;

        CComPtr<IUnknown> tmp1(IUnknownFactory::CreateInstance(clsid));
        if (!tmp1)
            return E_FAIL;

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

} // namespace ATL

struct SAFEARRAY {
    /** std::vector variant that also support non-owning pointers. */
    template <class T>
    class Buffer {
        public:
            Buffer(size_t size = 0) : m_size(size) {
                if (size > 0) {
                    m_ptr = new T[size];
                    m_owning = true;
                }
            }
            Buffer(const Buffer& other, bool deep_copy) : m_size(other.m_size) {
                if (deep_copy) {
                    m_ptr = new T[m_size];
                    m_owning = true;
                    memcpy(m_ptr, other.m_ptr, m_size);
                } else {
                    m_ptr = other.m_ptr;
                    m_owning = false;
                }
            }

            ~Buffer() {
                if (m_ptr && m_owning)
                    delete[] m_ptr;
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
                    new_ptr = new T[size];
                    m_owning = true;
                    memcpy(new_ptr, m_ptr, std::min(m_size, size));
                    for (size_t i = std::min(m_size, size); i < size; ++i)
                        new_ptr[i] = val;
                }
                // delete old buffer
                if (m_ptr && m_owning)
                    delete[] m_ptr;
                // commit changes
                m_ptr = new_ptr;
                m_size = size;
            }

            Buffer(const Buffer& other) = delete;
            Buffer(Buffer&&) = delete;
            Buffer& operator = (const Buffer&) = delete;
            Buffer& operator = (Buffer&&) = delete;

        private:
            size_t m_size = 0;
            T  *   m_ptr = nullptr;
            bool   m_owning = true;
    };

    enum TYPE {
        TYPE_EMPTY,
        TYPE_DATA,
        TYPE_STRINGS,
        TYPE_POINTERS,
    };

    SAFEARRAY (TYPE t) : type(t) {
        assert(t == TYPE_STRINGS || t == TYPE_POINTERS);
    }
    SAFEARRAY (unsigned int _elm_size, unsigned int count) : type(TYPE_DATA), data(_elm_size*count), elm_size(_elm_size) {
    }
    SAFEARRAY(const SAFEARRAY& other, bool deep_copy = true) : type(other.type), data(other.data, deep_copy), strings(other.strings), pointers(other.pointers), elm_size(other.elm_size) {
    }

    ~SAFEARRAY() {
    }

    SAFEARRAY () = delete;
    SAFEARRAY& operator = (const SAFEARRAY&) = delete;

    const TYPE                          type = TYPE_EMPTY; ///< \todo: Replace with std::variant when upgrading to C++17
    Buffer<unsigned char>               data;
    std::vector<ATL::CComBSTR>          strings;
    std::vector<ATL::CComPtr<IUnknown>> pointers;
    const unsigned int                  elm_size = 0;
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
    CComSafeArray (UINT size = 0) {
        if (size)
            m_ptr.reset(new SAFEARRAY(sizeof(T), size));
    }

    CComSafeArray (SAFEARRAY * obj) {
        if (obj)
            m_ptr.reset(new SAFEARRAY(*obj));
    }

    ~CComSafeArray () {
    }

    CComSafeArray (const CComSafeArray&) = delete;
    CComSafeArray (CComSafeArray&&) = default;
    CComSafeArray& operator = (const CComSafeArray&) = delete;
    CComSafeArray& operator = (CComSafeArray&&) = default;

    HRESULT Attach (SAFEARRAY * obj) {
        m_ptr.reset(obj);
        return S_OK;
    }
    SAFEARRAY* Detach () {
        return m_ptr.release();
    }

    operator SAFEARRAY* () {
        return m_ptr.get();
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
        assert(m_ptr);
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

    std::unique_ptr<SAFEARRAY> m_ptr;
};
// Template specializations. Implemented in cpp file.
template <> CComSafeArray<BSTR>::CComSafeArray (UINT size);
template <> CComSafeArray<IUnknown*>::CComSafeArray (UINT size);
template <> CComTypeWrapper<BSTR>::type& CComSafeArray<BSTR>::GetAt (int idx) const;
template <> CComTypeWrapper<IUnknown*>::type& CComSafeArray<IUnknown*>::GetAt (int idx) const;
template <> HRESULT                      CComSafeArray<BSTR>::Add (const typename CComTypeWrapper<BSTR>::type& t, BOOL copy);
template <> HRESULT                 CComSafeArray<IUnknown*>::Add (const typename CComTypeWrapper<IUnknown*>::type& t, BOOL copy);
template <> unsigned int CComSafeArray<BSTR>::GetCount () const;
template <> unsigned int CComSafeArray<IUnknown*>::GetCount () const;


#define SUCCEEDED(hr) (((HRESULT)(hr)) >= 0)
#define FAILED(hr)    (((HRESULT)(hr)) < 0)

// COM calling convention (use default on non-Windows)
#define STDMETHODCALLTYPE

#define ATL_NO_VTABLE 

// QueryInterface support macros
#define BEGIN_COM_MAP(CLASS)         HRESULT QueryInterface (const GUID & iid, /*out*/void **obj) override { \
                                           *obj = nullptr;

#define COM_INTERFACE_ENTRY(INTERFACE) if (iid == __uuidof(INTERFACE)) \
                                           *obj = static_cast<INTERFACE*>(this); \
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
class CComObjectRootEx {};

template <class T>
class CComObject {
public:
    static HRESULT CreateInstance (CComObject<T> ** arg) {
        assert(arg);
        assert(!*arg);

        *arg = new T;
        return S_OK;
    }
};

template <class T, const GUID* pclsid = nullptr>
class CComCoClass : public CComObject<T> {
};


} // namespace ATL

#ifndef _ATL_NO_AUTOMATIC_NAMESPACE
  using namespace ATL;
#endif
