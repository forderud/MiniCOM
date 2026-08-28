/* "App API" convenience functions.
   Copyright (c), GE HealthCare Ultrasound. */
#pragma once

#include <stdexcept>
#include <cassert>
#include <string>
#include <vector>
#include <cstdlib>


/** Converts unicode string to ASCII */
inline std::string ToAscii (const std::wstring& w_str) {
#ifdef _MSC_VER
#pragma warning(push)
#pragma warning(disable: 4996) // function or variable may be unsafe
#endif
    std::string s_str(w_str.size(), '\0');
    wcstombs(const_cast<char*>(s_str.data()), w_str.c_str(), s_str.size());
#ifdef _MSC_VER
#pragma warning(pop)
#endif
    return s_str;
}

/** Converts ASCII string to unicode */
inline std::wstring ToUnicode (const std::string& s_str) {
#ifdef _MSC_VER
#pragma warning(push)
#pragma warning(disable: 4996) // function or variable may be unsafe
#endif
    std::wstring w_str(s_str.size(), L'\0');
    mbstowcs(const_cast<wchar_t*>(w_str.data()), s_str.c_str(), w_str.size());
#ifdef _MSC_VER
#pragma warning(pop)
#endif
    return w_str;
}

#ifdef _WIN32
#include <comdef.h>  // for _com_error
#include <atlsafe.h> // for CComSafeArray
#include <atlcom.h>  // for CComObject
#include <atlbase.h> // for CComBSTR


/** Translate COM HRESULT failure into exceptions. */
inline void CHECK (HRESULT hr) {
    if (FAILED(hr)) {
        _com_error err(hr);
#ifdef _UNICODE
        const wchar_t * msg = err.ErrorMessage(); // weak ptr.
        throw std::runtime_error(ToAscii(msg));
#else
        const char * msg = err.ErrorMessage(); // weak ptr.
        throw std::runtime_error(msg);
#endif
    }
}


/* Disable BSTR caching to ease memory management debugging.
   REF: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/automat/setoanocache */
extern "C" void SetOaNoCache ();



/** RAII class for COM initialization. */
class ComInitialize {
public:
    ComInitialize (COINIT apartment /*= COINIT_MULTITHREADED*/) : m_initialized(false) {
        // REF: https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-coinitializeex
        HRESULT hr = CoInitializeEx(NULL, apartment);
        if (SUCCEEDED(hr))
            m_initialized = true;

#ifndef NDEBUG
        SetOaNoCache();
#endif
    }

    ~ComInitialize () {
        if (m_initialized)
            CoUninitialize();
    }

private:
    bool m_initialized; ///< must uninitialize in dtor
};

#else // _WIN32
#  include "NonWindows.hpp"

enum COINIT { 
  COINIT_APARTMENTTHREADED  = 0x2,
  COINIT_MULTITHREADED      = 0x0,
  COINIT_DISABLE_OLE1DDE    = 0x4,
  COINIT_SPEED_OVER_MEMORY  = 0x8
};

class ComInitialize {
public:
    ComInitialize(COINIT apartment) {
        (void)apartment;
    }
};

#endif // not _WIN32

/** Convenience function to create a locally implemented COM instance without the overhead of CoCreateInstance.
    The COM class does not need to be registred for construction to succeed. However, lack of registration can
    cause problems if transporting the class out-of-process. */
template <class T>
CComPtr<T> CreateLocalInstance () {
    // create an object (with ref. count zero)
    CComObject<T> * tmp = nullptr;
    CHECK(CComObject<T>::CreateInstance(&tmp));

    // move into smart-ptr (will incr. ref. count to one)
    return CComPtr<T>(static_cast<T*>(tmp));
}

/** Convenience converter to CComBSTR. */
inline CComBSTR ToCComBSTR (const std::wstring & str) {
    if (str.empty())
        return CComBSTR(L""); // special handling to avoid nullptr m_str

    return CComBSTR(static_cast<int>(str.size()), str.data());
}

/** Convenience converter from CComBSTR. */
inline std::wstring ToWString(const CComBSTR & str) {
    unsigned int str_len = str.Length();
    if (!str_len)
        return std::wstring();

    return std::wstring(str.m_str, str_len);
}

/** Convert SafeArray to a std::vector. */
template <class T>
std::vector<T> ConvertToVector (const CComSafeArray<T> & input) {
    std::vector<T> result(input.GetCount());
    if (result.size() > 0) {
        const T * in_ptr = &input.GetAt(0); // will fail if empty
        for (size_t i = 0; i < result.size(); ++i)
            result[i] = in_ptr[i];
    }
    return result;
}

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
