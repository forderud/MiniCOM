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
#include <string.h> // for memcmp
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
static GUID hold_uuidof () = delete; // fail build if performing __uuidof() on unregistered interface
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
typedef unsigned short VARTYPE; ///< VARENUM value identifying a type
typedef wchar_t*       BSTR;    ///< zero terminated text string prefixed by a 32bit byte length (see SysAllocString)
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
#define REGDB_E_CLASSNOTREG static_cast<int32_t>(0x80040154L)
#define CO_E_CLASSSTRING    static_cast<int32_t>(0x800401F3L)
#define DISP_E_BADINDEX      static_cast<int32_t>(0x8002000BL)
#define DISP_E_ARRAYISLOCKED static_cast<int32_t>(0x8002000DL)

/** Subset of the VARENUM type codes from wtypes.h. Used to tag SAFEARRAY element types. */
enum VARENUM {
  VT_EMPTY   = 0,
  VT_I2      = 2,
  VT_I4      = 3,
  VT_R4      = 4,
  VT_R8      = 5,
  VT_BSTR    = 8,
  VT_UNKNOWN = 13,
  VT_I1      = 16,
  VT_UI1     = 17,
  VT_UI2     = 18,
  VT_UI4     = 19,
  VT_I8      = 20,
  VT_UI8     = 21
};


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
    case DISP_E_BADINDEX:      return "DISP_E_BADINDEX";
    case DISP_E_ARRAYISLOCKED: return "DISP_E_ARRAYISLOCKED";
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

#ifdef _UNICODE
        size_t len = strlen(str);
        m_buffer.resize(len, L'\0');
        mbstowcs(const_cast<wchar_t*>(m_buffer.data()), str, len);
#else
        m_buffer = str;
#endif
    }

    HRESULT Error() const noexcept {
        return m_hr;
    }

#ifdef _UNICODE
    const wchar_t* ErrorMessage() const noexcept {
#else
    const char* ErrorMessage() const noexcept {
#endif
        return m_buffer.c_str();
    }

private:
    HRESULT      m_hr = E_FAIL;
#ifdef _UNICODE
    std::wstring m_buffer;
#else
    std::string  m_buffer;
#endif
};


/* BSTR memory layout, same as on Windows:
     [byte length (32bit)] [characters...] [zero terminator]
   The BSTR pointer points to the first character, with the byte length stored in the 4 bytes before it.
   The length excludes the zero terminator and is sizeof(wchar_t) per character, which is 4 on Linux & MacOS
   (UTF-32) as opposed to 2 on Windows (UTF-16). A BSTR must be freed with SysFreeString, not free(). */
extern "C" { // to avoid name mangling
/** Allocate a BSTR copy of a zero terminated string. Returns nullptr if str is nullptr. */
BSTR SysAllocString (const wchar_t* str);
/** Allocate a BSTR of "len" characters, initialized from "str" if non-null. Can contain embedded zeros. */
BSTR SysAllocStringLen (const wchar_t* str, unsigned int len);
/** Free a BSTR allocated with SysAllocString or SysAllocStringLen. Accepts nullptr. */
void SysFreeString (BSTR str);
/** Returns the BSTR length in characters, excluding the zero terminator (0 for nullptr). */
unsigned int SysStringLen (BSTR str);
/** Returns the BSTR length in bytes, excluding the zero terminator (0 for nullptr). */
unsigned int SysStringByteLen (BSTR str);

/** COM task memory allocator, used for [out] arguments that the callee allocates and the caller frees. */
void* CoTaskMemAlloc (size_t size);
void* CoTaskMemRealloc (void* ptr, size_t size);
void  CoTaskMemFree (void* ptr);
} // extern "C"


/** API-compatible subset of the Microsoft _bstr_t class documented on https://docs.microsoft.com/en-us/cpp/cpp/bstr-t-class */
class _bstr_t {
public:
    _bstr_t() noexcept = default;

    _bstr_t(const _bstr_t& s) noexcept {
        AssignCopy(s.m_str);
    }
    _bstr_t(const wchar_t* s) {
        Assign(s);
    }
    _bstr_t(BSTR s, bool copy) {
        if (copy)
            AssignCopy(s);
        else
            m_str = s; // attach to BSTR
    }

    ~_bstr_t() noexcept {
        Clear();
    }

    _bstr_t& operator=(const _bstr_t& s) noexcept {
        AssignCopy(s.m_str);
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
        return SysStringLen(m_str);
    }

    void Assign(const wchar_t* s) {
        if (s == m_str)
            return; // self-assignment

        Clear();

        if (s)
            m_str = SysAllocString(s);
    }

    BSTR* GetAddress() {
        Clear();
        return &m_str;
    }

    /** Take over ownership of a BSTR (does not copy). */
    void Attach(BSTR s) {
        Clear();
        m_str = s;
    }

    /** Release ownership of the BSTR. Must be freed with SysFreeString unless re-attached. */
    BSTR Detach() {
        BSTR tmp = m_str;
        m_str = nullptr;
        return tmp;
    }

private:
    /** Copy a BSTR including its length, so that embedded zeros are kept. */
    void AssignCopy(BSTR s) {
        if (s == m_str)
            return; // self-assignment

        Clear();

        if (s)
            m_str = SysAllocStringLen(s, SysStringLen(s));
    }

    void Clear() {
        if (!m_str)
            return;

        SysFreeString(m_str);
        m_str = nullptr;
    }

    BSTR m_str = nullptr;
};
static_assert(sizeof(_bstr_t) == sizeof(wchar_t*), "_bstr_t size mismatch");


namespace ATL {

class CComBSTR {
public:
    CComBSTR () {
    }
    CComBSTR (const wchar_t* str) {
        if (str)
            m_str = SysAllocString(str);
    }
    /** Copy "size" characters from str (nullptr if size is zero), like ATL. */
    CComBSTR (int size, const wchar_t* str) {
        assert(size >= 0);
        if (size > 0)
            m_str = SysAllocStringLen(str, static_cast<unsigned int>(size));
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
        return SysStringLen(m_str);
    }

    /** Returns string length in bytes excluding null termination. */
    unsigned int ByteLength () const {
        return SysStringByteLen(m_str);
    }

    operator BSTR () const {
        return m_str;
    }

    BSTR* operator & () {
        return &m_str;
    }

    /** Take over ownership of a BSTR (does not copy). */
    void Attach(BSTR s) noexcept {
        if (s == m_str)
            return;

        Empty();
        m_str = s;
    }

    /** Release ownership of the BSTR. Must be freed with SysFreeString unless re-attached. */
    BSTR Detach () {
        BSTR tmp = m_str;
        m_str = nullptr;
        return tmp;
    }

    /** Returns a BSTR copy that the caller must free with SysFreeString. */
    BSTR Copy () const {
        if (!m_str)
            return nullptr;

        return SysAllocStringLen(m_str, SysStringLen(m_str));
    }

    void Empty() {
        if (!m_str)
            return;

        SysFreeString(m_str);
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

    BSTR m_str = nullptr;
};
static_assert(sizeof(CComBSTR) == sizeof(BSTR), "CComBSTR size mismatch");

template<typename T>
class CComPtr;
} // namespace ATL

template<typename T>
class _com_ptr_t;


extern "C" {
// interface ID values for well-known interfaces
static constexpr GUID IID_IUnknown       = {0x00000000,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static constexpr GUID IID_IMessageFilter = {0x00000016,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
} // extern "C"

#if !defined(CINTERFACE)

/** IUnknown base-class for non-Windows platforms. */
struct IUnknown {
    /** Cast method. */
    virtual HRESULT QueryInterface (const GUID& iid, /*out*/void** obj) = 0;

    /** Typed cast method. */
    template <class Q>
    HRESULT QueryInterface (Q** obj) {
        return QueryInterface(__uuidof(Q), reinterpret_cast<void**>(obj));
    }

    /** Reference counting for lifetime management. */
    virtual ULONG AddRef () = 0;
    virtual ULONG Release () = 0;
};
DEFINE_UUIDOF(IUnknown)

#else // defined(CINTERFACE)

struct IUnknown; // forward declaration

typedef struct IUnknownVtbl {
    HRESULT (*QueryInterface)(IUnknown* This, const GUID& iid, void** obj);
    ULONG   (*AddRef)(IUnknown* This);
    ULONG   (*Release)(IUnknown* This);
} IUnknownVtbl;

struct IUnknown {
    struct IUnknownVtbl* lpVtbl;
};

#endif


/** Resolve COM class CLSID based on "[<Program>.]<Component>[.<Version>]" ProgID string. */
extern "C" // to avoid name mangling
HRESULT CLSIDFromProgID (const wchar_t* ProgID, /*out*/GUID* clsid);

/** Create COM class based on CLSID. */
extern "C" // to avoid name mangling
HRESULT CoCreateInstance (const GUID& clsid, IUnknown* outer, DWORD context, const GUID& iid, /*out*/void** result);


// error handler required by generated wrapper API headers
inline void _com_issue_errorex(HRESULT hr, IUnknown*, const IID &) {
    throw _com_error(hr);
}


#if !defined(CINTERFACE)


template <class BASE>
class CComObject final : public BASE {
public:
    ULONG AddRef () override {
        return BASE::InternalAddRef();
    }
    ULONG Release () override {
        ULONG ref = BASE::InternalRelease();
        if (!ref)
            delete this;

        return ref;
    }

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
class CComAggObject final : public IUnknown {
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
        if (!ref) {
            delete this;
        }
        return ref;
    }
    HRESULT QueryInterface (const GUID & iid, /*out*/void **obj) override {
        if (!obj)
            return E_POINTER;
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
    friend HRESULT CLSIDFromProgID (const wchar_t* ProgID, /*out*/GUID* clsid);
    friend HRESULT CoCreateInstance (const GUID& clsid, IUnknown* outer, DWORD context, const GUID& iid, /*out*/void** result);

private:
    typedef HRESULT(*Factory)(IUnknown*, IUnknown**);
    
    struct Entry {
        GUID          clsid{};
        ATL::CComBSTR name;
        Factory       factory = nullptr;
    };
    
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

            return tmp->QueryInterface(__uuidof(IUnknown), reinterpret_cast<void**>(obj)); // incr. ref-count to one
        } else {
            // create an object (with ref. count zero)
            CComObject<CLS> * tmp = nullptr;
            HRESULT hr = CComObject<CLS>::CreateInstance(&tmp);
            if (FAILED(hr))
                return hr;

            return tmp->QueryInterface(__uuidof(IUnknown), reinterpret_cast<void**>(obj)); // incr. ref-count to one
        }
    }

    static Buffer<Entry> & Factories ();
};

#define OBJECT_ENTRY_AUTO(clsid, cls) \
    __attribute__((weak)) __attribute__((used)) const char* tmp_factory_##cls = IUnknownFactory::RegisterClass<cls>(clsid, #cls);


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

    bool operator==(const _com_ptr_t& p) const {
        if (m_ptr == p.m_ptr)
            return true;

        return CompareUnknown(p.m_ptr) == 0;
    }

    template<typename U, ::std::enable_if_t<!::std::is_same<U, T>::value, int> = 0>
    bool operator==(const _com_ptr_t<U>& other) const {
        U* other_ptr = other; // cannot access other.m_ptr member directly, since it's private
        return CompareUnknown(other_ptr) == 0;
    }

    bool operator==(T* p) const {
        if (m_ptr == p)
            return true;
        return CompareUnknown(p) == 0;
    }

    template<typename U, ::std::enable_if_t<!::std::is_same<U, T>::value, int> = 0>
    bool operator==(U* p) const {
        return CompareUnknown(p) == 0;
    }

    template<typename U>
    friend bool operator== (U* left, const _com_ptr_t& right) {
        static_assert(std::is_base_of<IUnknown, U>::value, "_com_ptr_t::CompareUnknown: U must inherit from IUnknown");
        return right == left;
    }

    HRESULT CreateInstance (const GUID& clsid, IUnknown* outer = nullptr, DWORD context = CLSCTX_ALL) noexcept {
        _com_ptr_t tmp;
        HRESULT hr = ::CoCreateInstance(clsid, outer, context, __uuidof(T), (void**)&tmp);
        if (FAILED(hr))
            return hr;

        Swap(tmp);
        return S_OK;
    }

    HRESULT CreateInstance(const wchar_t* name, IUnknown* outer = nullptr, DWORD context = CLSCTX_ALL) noexcept {
        if (!name)
            return E_INVALIDARG;
        
        GUID clsid{};
        HRESULT hr = CLSIDFromProgID(name, &clsid);
        if (FAILED(hr))
            return hr;

        _com_ptr_t tmp;
        hr = ::CoCreateInstance(clsid, outer, context, __uuidof(T), (void**)&tmp);
        if (FAILED(hr))
            return hr;

        Swap(tmp);
        return S_OK;
    }

private:
    T * m_ptr = nullptr;

    void Swap (_com_ptr_t & other) {
        T* tmp = m_ptr;
        m_ptr = other.m_ptr;
        other.m_ptr = tmp;
    }

    template<typename U>
    ptrdiff_t CompareUnknown(U * p) const {
        static_assert(std::is_base_of<IUnknown, U>::value, "_com_ptr_t::CompareUnknown: U must inherit from IUnknown");

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
        GUID clsid{};
        HRESULT hr = CLSIDFromProgID(name.c_str(), &clsid);
        if (FAILED(hr))
            return hr;

        CComPtr tmp;
        hr = ::CoCreateInstance(clsid, outer, context, __uuidof(T), (void**)&tmp); // RefCount=1
        if (FAILED(hr))
            return hr;

        Swap(tmp);
        return S_OK;
    }

    HRESULT CoCreateInstance (GUID clsid, IUnknown* outer = NULL, DWORD context = CLSCTX_ALL) {
        CComPtr tmp;
        HRESULT hr = ::CoCreateInstance(clsid, outer, context, __uuidof(T), (void**)&tmp); // RefCount=1
        if (FAILED(hr))
            return hr;

        Swap(tmp);
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


/** Taken from oaidl.h. */
struct SAFEARRAYBOUND {
    ULONG cElements; ///< number of elements
    LONG  lLbound;   ///< index of the first element
};

/** Taken from oaidl.h. Same memory layout as on Windows: the elements are stored contiguously in pvData, with
    BSTR and IUnknown* elements as plain pointers. The descriptor is preceded by 16 hidden bytes, with the element
    VARTYPE stored in the 4 bytes before the descriptor (FADF_HAVEVARTYPE), like on Windows.
    Only one-dimensional arrays are supported. */
struct SAFEARRAY {
    USHORT cDims;      ///< number of dimensions (always 1)
    USHORT fFeatures;  ///< FADF_* flags
    ULONG  cbElements; ///< element size in bytes
    ULONG  cLocks;     ///< lock count (incremented by SafeArrayAccessData)
    void*  pvData;     ///< element data
    SAFEARRAYBOUND rgsabound[1]; ///< one bound per dimension
};
static_assert(sizeof(SAFEARRAY) == 16 + 2*sizeof(void*), "SAFEARRAY not packed");

// SAFEARRAY fFeatures flags (from oaidl.h)
#define FADF_HAVEVARTYPE 0x0080 ///< element VARTYPE is stored before the descriptor
#define FADF_BSTR        0x0100 ///< array of BSTRs (freed with SysFreeString when destroyed)
#define FADF_UNKNOWN     0x0200 ///< array of IUnknown* (released when destroyed)

/* One-dimensional subset of the SAFEARRAY API, with the same signatures as on Windows. */
extern "C" { // to avoid name mangling
/** Create a zero-initialized array of the given element type. Only cDims=1 is supported. Returns nullptr on failure. */
SAFEARRAY* SafeArrayCreate (VARTYPE vt, UINT cDims, SAFEARRAYBOUND* rgsabound);
SAFEARRAY* SafeArrayCreateVector (VARTYPE vt, LONG lLbound, ULONG cElements);
/** Free the elements (BSTRs & IUnknown* included), the data and the descriptor. Fails if locked. */
HRESULT SafeArrayDestroy (SAFEARRAY* psa);
/** Get the element type. Fails for arrays without FADF_HAVEVARTYPE. */
HRESULT SafeArrayGetVartype (SAFEARRAY* psa, /*out*/VARTYPE* pvt);
UINT    SafeArrayGetDim (SAFEARRAY* psa);
UINT    SafeArrayGetElemsize (SAFEARRAY* psa);
HRESULT SafeArrayGetLBound (SAFEARRAY* psa, UINT nDim, /*out*/LONG* plLbound);
HRESULT SafeArrayGetUBound (SAFEARRAY* psa, UINT nDim, /*out*/LONG* plUbound);
/** Lock the array and return a pointer to its data. Call SafeArrayUnaccessData when done. */
HRESULT SafeArrayAccessData (SAFEARRAY* psa, /*out*/void** ppvData);
HRESULT SafeArrayUnaccessData (SAFEARRAY* psa);
} // extern "C"

/** Internal function used by CComSafeArray, to also support element types without a VARTYPE (vt=VT_EMPTY). */
SAFEARRAY* InternalSafeArrayCreate (VARTYPE vt, ULONG elm_size, ULONG count);


namespace ATL {

/** Maps SAFEARRAY element types to their VARTYPE and the class that wraps the elements in-place. */
template <typename T>
struct CComTypeWrapper {
    typedef T type; // default mapping
    static constexpr VARTYPE vt = VT_EMPTY; // no VARTYPE for other types (e.g. structs), so SafeArrayGetVartype fails
};
#define DEFINE_COM_TYPE_WRAPPER(T, WRAPPER, VT) \
    template <> \
    struct CComTypeWrapper<T> { \
        typedef WRAPPER type; \
        static constexpr VARTYPE vt = VT; \
    };
DEFINE_COM_TYPE_WRAPPER(char,               char,               VT_I1)
DEFINE_COM_TYPE_WRAPPER(unsigned char,      unsigned char,      VT_UI1)
DEFINE_COM_TYPE_WRAPPER(short,              short,              VT_I2)
DEFINE_COM_TYPE_WRAPPER(unsigned short,     unsigned short,     VT_UI2)
DEFINE_COM_TYPE_WRAPPER(int,                int,                VT_I4)
DEFINE_COM_TYPE_WRAPPER(unsigned int,       unsigned int,       VT_UI4)
DEFINE_COM_TYPE_WRAPPER(long long,          long long,          VT_I8)
DEFINE_COM_TYPE_WRAPPER(unsigned long long, unsigned long long, VT_UI8)
DEFINE_COM_TYPE_WRAPPER(float,              float,              VT_R4)
DEFINE_COM_TYPE_WRAPPER(double,             double,             VT_R8)
DEFINE_COM_TYPE_WRAPPER(BSTR,               CComBSTR,           VT_BSTR)    // map BSTR/wchar_t* to CComBSTR
DEFINE_COM_TYPE_WRAPPER(IUnknown*,          CComPtr<IUnknown>,  VT_UNKNOWN) // wrap IUnknown* in CComPtr
#undef DEFINE_COM_TYPE_WRAPPER
static_assert(sizeof(CComPtr<IUnknown>) == sizeof(IUnknown*), "CComPtr size mismatch");

template <class T>
struct CComSafeArray {
    typedef typename CComTypeWrapper<T>::type ELM; ///< in-place element type

    CComSafeArray () {
    }

    CComSafeArray (UINT size) {
        m_ptr = InternalSafeArrayCreate(CComTypeWrapper<T>::vt, sizeof(T), size);
        assert(m_ptr);
        InternalConstruct(0, size);
    }

    CComSafeArray (SAFEARRAY * obj) {
        if (obj) {
            assert(obj->cbElements == sizeof(T));
            unsigned int count = InternalElementCount(obj);
            m_ptr = InternalSafeArrayCreate(CComTypeWrapper<T>::vt, sizeof(T), count);
            assert(m_ptr);
            InternalConstruct(0, count);
            for (unsigned int i = 0; i < count; i++)
                GetAt(i) = InternalDataPointer(obj)[i]; // deep copy
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
            if constexpr (CComTypeWrapper<T>::vt == VT_EMPTY) {
                for (unsigned int i = 0; i < GetCount(); i++)
                    GetAt(i).~ELM(); // SafeArrayDestroy only knows how to destroy BSTR & IUnknown* elements
            }
            HRESULT hr = SafeArrayDestroy(m_ptr);
            assert(SUCCEEDED(hr) && "CComSafeArray::Destroy failed (array locked?)");
            if (FAILED(hr))
                return hr;
            m_ptr = nullptr;
        }
        return S_OK;
    }

    /** Take over ownership of an existing array. */
    HRESULT Attach (SAFEARRAY * obj) {
        assert(obj);
        assert(obj->cbElements == sizeof(T));
        Destroy();
        m_ptr = obj;
        return S_OK;
    }
    /** Release ownership of the array. Must be destroyed with SafeArrayDestroy unless re-attached. */
    SAFEARRAY* Detach () {
        SAFEARRAY* tmp = m_ptr;
        m_ptr = nullptr;
        return tmp;
    }

    operator SAFEARRAY* () {
        return m_ptr;
    }

    /** Returns a reference to the element. BSTR & IUnknown* elements are returned as CComBSTR & CComPtr<IUnknown> references into the array. */
    ELM& GetAt (int idx) const {
        assert(m_ptr);
        assert((idx >= 0) && (static_cast<unsigned int>(idx) < GetCount()));
        return InternalDataPointer(m_ptr)[idx];
    }

    ELM& operator [] (int idx) {
        return GetAt(idx);
    }

    /** Set an element. BSTR & IUnknown* elements are copied/AddRef'ed. */
    HRESULT SetAt (int idx, const T& val, bool copy = true) {
        (void)copy; // mute unreferenced argument warning
        if constexpr (std::is_same_v<T, IUnknown*>)
            assert(val);

        GetAt(idx) = val;
        return S_OK;
    }

    /** Append an element. BSTR & IUnknown* elements are copied/AddRef'ed. */
    HRESULT Add (const ELM& t, BOOL copy = true) {
        (void)copy; // mute unreferenced argument warning

        if (!m_ptr)
            m_ptr = InternalSafeArrayCreate(CComTypeWrapper<T>::vt, sizeof(T), 0); // lazy initialization

        assert(sizeof(T) == m_ptr->cbElements);
        const unsigned int prev_count = GetCount();
        void* data = realloc(m_ptr->pvData, (prev_count + 1)*sizeof(T));
        if (!data)
            return E_OUTOFMEMORY;

        m_ptr->pvData = data;
        m_ptr->rgsabound[0].cElements = prev_count + 1;
        InternalConstruct(prev_count, prev_count + 1);
        GetAt(prev_count) = t;
        return S_OK;
    }

    unsigned int GetCount () const {
        assert(m_ptr);
        return m_ptr->rgsabound[0].cElements;
    }

    /** Internal function. Do NOT call unless you know what you're doing. */
    static ELM* InternalDataPointer(SAFEARRAY* obj) {
        return static_cast<ELM*>(obj->pvData);
    }

    /** Internal function. Do NOT call unless you know what you're doing. */
    static unsigned int InternalElementCount(SAFEARRAY* obj) {
        return obj->rgsabound[0].cElements;
    }

    SAFEARRAY* m_ptr = nullptr;

private:
    /** Default-construct the elements in the [first, last) range (the memory is already zero-initialized). */
    void InternalConstruct (unsigned int first, unsigned int last) {
        if constexpr (!std::is_trivially_default_constructible_v<ELM>) {
            for (unsigned int i = first; i < last; i++)
                new (&InternalDataPointer(m_ptr)[i]) ELM();
        } else {
            (void)first; (void)last; // mute unreferenced argument warnings
        }
    }
};
static_assert(sizeof(CComSafeArray<BYTE>) == sizeof(SAFEARRAY*), "CComSafeArray size mismatch");


// COM calling convention (use default on non-Windows)
#define STDMETHODCALLTYPE

#define ATL_NO_VTABLE

// QueryInterface support macros
#define BEGIN_COM_MAP(CLASS)         HRESULT QueryInterface (const GUID & iid, /*out*/void **obj) override { \
                                           static_assert(std::is_same_v<CLASS, std::remove_pointer_t<decltype(this)>>, \
                                               "Argument to BEGIN_COM_MAP doesn't match name of surrounding class."); \
                                           if (!obj) \
                                               return E_POINTER; \
                                           *obj = nullptr; \
                                           IUnknown* this_unknown = nullptr;
#define COM_INTERFACE_ENTRY(INTERFACE) if (!this_unknown) \
                                           this_unknown = static_cast<IUnknown*>(static_cast<INTERFACE*>(this)); \
                                       if (iid == __uuidof(INTERFACE)) { \
                                           *obj = static_cast<INTERFACE*>(this); \
                                           AddRef(); \
                                           return S_OK; \
                                       }
#define COM_INTERFACE_ENTRY_AGGREGATE(INTERFACE, punk) \
                                       if (iid == __uuidof(INTERFACE)) { \
                                           *obj = static_cast<INTERFACE*>(&punk->m_contained); \
                                           AddRef(); \
                                           return S_OK; \
                                       }
#define END_COM_MAP()                  if (iid == __uuidof(IUnknown)) { \
                                           assert(this_unknown && "COM_INTERFACE_ENTRY() entries missing"); \
                                           *obj = this_unknown; \
                                           AddRef(); \
                                           return S_OK; \
                                       } \
                                       return E_NOINTERFACE; \
                                     } \
                                     ULONG AddRef () override = 0; \
                                     ULONG Release () override = 0;

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

    // reference-counting used by CComObject
    ULONG InternalAddRef() {
        assert((m_ref < 0xFFFF) && "IUnknown::AddRef negative ref count.");
        return ++m_ref;
    }
    ULONG InternalRelease() {
        ULONG ref = --m_ref;
        assert((m_ref < 0xFFFF) && "IUnknown::Release negative ref count.");
        return ref;
    }

protected:
    std::atomic<ULONG> m_ref {0};
};

template <class T, const GUID* pclsid = nullptr>
class CComCoClass {
};


} // namespace ATL

#ifndef _ATL_NO_AUTOMATIC_NAMESPACE
  using namespace ATL;
#endif

#endif // !defined(CINTERFACE)
