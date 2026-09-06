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
static GUID hold_uuidof () = delete; // fail build if performing __uuidof() on unregistered interface
#define DEFINE_UUIDOF_ID(Q, IID) template<> inline GUID hold_uuidof<Q>() { return IID; }
#define DEFINE_UUIDOF(Q) DEFINE_UUIDOF_ID(Q, IID_##Q)
#define __uuidof(Q) hold_uuidof<Q>()

typedef GUID           IID;

// GUID reference type, matching guiddef.h
#define REFIID const IID &

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
#define REGDB_E_CLASSNOTREG static_cast<int32_t>(0x80040154L)
#define CO_E_CLASSSTRING    static_cast<int32_t>(0x800401F3L)


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

// COM calling convention (use default on non-Windows)
#define STDMETHODCALLTYPE

// vtable pointer qualifier and layout markers, matching rpcndr.h
#define CONST_VTBL const
#define BEGIN_INTERFACE
#define END_INTERFACE

extern "C" {
// interface ID values for well-known interfaces
static constexpr GUID IID_IUnknown       = {0x00000000,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static constexpr GUID IID_IMessageFilter = {0x00000016,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};

#if defined(__cplusplus) && !defined(CINTERFACE)
/** IUnknown base-class for non-Windows platforms. */
struct IUnknown {
    /** Cast method. */
    virtual HRESULT QueryInterface (const GUID & iid, /*[out]*/void **obj) = 0;

    /** Reference counting for lifetime management. */
    virtual ULONG AddRef () = 0;
    virtual ULONG Release () = 0;
};
#else /* C style interface */
struct IUnknown;

typedef struct IUnknownVtbl {
    BEGIN_INTERFACE
    HRESULT (STDMETHODCALLTYPE *QueryInterface) (IUnknown* This, REFIID riid, void** ppvObject);
    ULONG (STDMETHODCALLTYPE *AddRef) (IUnknown* This);
    ULONG (STDMETHODCALLTYPE *Release) (IUnknown* This);
    END_INTERFACE
} IUnknownVtbl;

struct IUnknown {
    CONST_VTBL struct IUnknownVtbl* lpVtbl;
};
#endif
} // extern "C"
DEFINE_UUIDOF(IUnknown)

/** Resolve COM class CLSID based on "[<Program>.]<Component>[.<Version>]" ProgID string. */
extern "C" // to avoid name mangling
HRESULT CLSIDFromProgID (const wchar_t* ProgID, /*out*/GUID* clsid);

/** Create COM class based on CLSID. */
extern "C" // to avoid name mangling
HRESULT CoCreateInstance (const GUID& clsid, IUnknown* outer, DWORD context, const GUID& iid, /*out*/void** result);
