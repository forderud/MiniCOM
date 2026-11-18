#include "NonWindows.hpp"


__attribute__((visibility("default")))
BSTR SysAllocStringLen (const wchar_t* str, unsigned int len) {
    if (len > (UINT32_MAX - sizeof(wchar_t))/sizeof(wchar_t))
        return nullptr; // byte length doesn't fit in the 32bit prefix

    const size_t byte_len = static_cast<size_t>(len)*sizeof(wchar_t);

    // byte length prefix, followed by the characters and a zero terminator
    auto* ptr = static_cast<unsigned char*>(malloc(sizeof(uint32_t) + byte_len + sizeof(wchar_t)));
    if (!ptr)
        return nullptr;

    *reinterpret_cast<uint32_t*>(ptr) = static_cast<uint32_t>(byte_len);
    BSTR result = reinterpret_cast<BSTR>(ptr + sizeof(uint32_t));
    if (str)
        memcpy(result, str, byte_len); // might contain embedded zeros
    else
        memset(result, 0, byte_len);
    result[len] = L'\0';
    return result;
}

__attribute__((visibility("default")))
BSTR SysAllocString (const wchar_t* str) {
    if (!str)
        return nullptr;

    return SysAllocStringLen(str, static_cast<unsigned int>(wcslen(str)));
}

__attribute__((visibility("default")))
void SysFreeString (BSTR str) {
    if (!str)
        return;

    free(reinterpret_cast<unsigned char*>(str) - sizeof(uint32_t));
}

__attribute__((visibility("default")))
unsigned int SysStringByteLen (BSTR str) {
    if (!str)
        return 0;

    return *reinterpret_cast<uint32_t*>(reinterpret_cast<unsigned char*>(str) - sizeof(uint32_t));
}

__attribute__((visibility("default")))
unsigned int SysStringLen (BSTR str) {
    return SysStringByteLen(str)/sizeof(wchar_t);
}


__attribute__((visibility("default")))
void* CoTaskMemAlloc (size_t size) {
    return malloc(size);
}

__attribute__((visibility("default")))
void* CoTaskMemRealloc (void* ptr, size_t size) {
    return realloc(ptr, size);
}

__attribute__((visibility("default")))
void CoTaskMemFree (void* ptr) {
    free(ptr);
}


/** Resolve COM class CLSID based on "[<Program>.]<Component>[.<Version>]" ProgID string. */
__attribute__((visibility("default")))
HRESULT CLSIDFromProgID (const wchar_t* ProgID, /*out*/GUID* clsid) {
    if (!ProgID || !clsid)
        return E_POINTER;
    
    std::wstring class_name = ProgID;

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

    for (size_t i = 0; i < IUnknownFactory::Factories().size(); i++) {
        const auto & elm = IUnknownFactory::Factories()[i];
        if (elm.name == ATL::CComBSTR(class_name.c_str())) {
            *clsid = elm.clsid;
            return S_OK;
        }
    }

    std::wcerr << L"CLSIDFromProgID error: Unknown class " << class_name << std::endl;
    return CO_E_CLASSSTRING;
}

/** Create COM class based on CLSID. */
__attribute__((visibility("default")))
HRESULT CoCreateInstance (const GUID& clsid, IUnknown* outer, DWORD context, const GUID& iid, /*out*/void** result) {
    (void)context;

    if (!result)
        return E_POINTER;
    *result = nullptr;

    if (outer && !(iid == IID_IUnknown))
        return E_INVALIDARG;

    for (size_t i = 0; i < IUnknownFactory::Factories().size(); i++) {
        const auto & elm = IUnknownFactory::Factories()[i];
        if (elm.clsid == clsid) {
            IUnknown* obj = nullptr;
            HRESULT hr = elm.factory(outer, &obj); // RefCount=1
            assert(hr == S_OK);
            if (FAILED(hr))
                return hr;

            // cast to requested interface
            assert(obj);
            hr = obj->QueryInterface(iid, result);
            obj->Release();
            obj = nullptr;
            return hr;
        }
    }

    char guid_str[39] = {};
    snprintf(guid_str, sizeof(guid_str), "{%08x-%04x-%04x-%02x%02x-%02x%02x%02x%02x%02x%02x}",
        clsid.Data1, clsid.Data2, clsid.Data3,
        clsid.Data4[0], clsid.Data4[1], clsid.Data4[2], clsid.Data4[3],
        clsid.Data4[4], clsid.Data4[5], clsid.Data4[6], clsid.Data4[7]);

    std::cerr << "CoCreateInstance error: Unknown clsid " << guid_str << std::endl;
    return REGDB_E_CLASSNOTREG;
}


__attribute__((visibility("default")))
Buffer<IUnknownFactory::Entry> & IUnknownFactory::Factories () {
    static Buffer<Entry> s_factory;
    return s_factory;
}


template <> __attribute__((visibility("default")))
ATL::CComSafeArray<BSTR>::CComSafeArray (UINT size) {
    m_ptr = SAFEARRAY::Create(SAFEARRAY::TYPE_STRINGS);
    m_ptr->strings.resize(size);
}
template <> __attribute__((visibility("default")))
ATL::CComSafeArray<IUnknown*>::CComSafeArray (UINT size) {
    m_ptr = SAFEARRAY::Create(SAFEARRAY::TYPE_POINTERS);
    m_ptr->pointers.resize(size);
}

template <> __attribute__((visibility("default")))
ATL::CComTypeWrapper<BSTR>::type& CComSafeArray<BSTR>::GetAt (int idx) const {
    assert(m_ptr);
    assert(m_ptr->type == SAFEARRAY::TYPE_STRINGS);
    return m_ptr->strings[idx];
}
template <> __attribute__((visibility("default")))
ATL::CComTypeWrapper<IUnknown*>::type& CComSafeArray<IUnknown*>::GetAt (int idx) const {
    assert(m_ptr);
    assert(m_ptr->type == SAFEARRAY::TYPE_POINTERS);
    return m_ptr->pointers[idx];
}

template <> __attribute__((visibility("default")))
HRESULT ATL::CComSafeArray<IUnknown*>::SetAt(int idx, IUnknown* const& val, [[maybe_unused]] bool copy) {
    assert(val);
    assert(static_cast<size_t>(idx) < m_ptr->pointers.size());

    if (static_cast<size_t>(idx) >= m_ptr->pointers.size())
        return E_INVALIDARG;

    assert(m_ptr);
    assert(m_ptr->type == SAFEARRAY::TYPE_POINTERS);
    m_ptr->pointers[idx] = val;
    return S_OK;
}

template <> __attribute__((visibility("default")))
HRESULT ATL::CComSafeArray<BSTR>::Add (const typename CComTypeWrapper<BSTR>::type& t, BOOL copy) {
    (void)copy; // mute unreferenced argument warning

    if (!m_ptr)
        m_ptr = SAFEARRAY::Create(SAFEARRAY::TYPE_STRINGS); // lazy initialization

    assert(m_ptr->type == SAFEARRAY::TYPE_STRINGS);
    const size_t prev_size = m_ptr->strings.size();
    m_ptr->strings.resize(prev_size + 1, t);
    return S_OK;
}
template <> __attribute__((visibility("default")))
HRESULT ATL::CComSafeArray<IUnknown*>::Add (const typename CComTypeWrapper<IUnknown*>::type& t, BOOL copy) {
    (void)copy; // mute unreferenced argument warning

    if (!m_ptr)
        m_ptr = SAFEARRAY::Create(SAFEARRAY::TYPE_POINTERS); // lazy initialization

    assert(m_ptr->type == SAFEARRAY::TYPE_POINTERS);
    const size_t prev_size = m_ptr->pointers.size();
    m_ptr->pointers.resize(prev_size + 1, t);
    return S_OK;
}

template <> __attribute__((visibility("default")))
unsigned int ATL::CComSafeArray<BSTR>::GetCount () const {
    assert(m_ptr);
    assert(m_ptr->type == SAFEARRAY::TYPE_STRINGS);
    return static_cast<unsigned int>(m_ptr->strings.size());
}
template <> __attribute__((visibility("default")))
unsigned int ATL::CComSafeArray<IUnknown*>::GetCount () const {
    assert(m_ptr);
    assert(m_ptr->type == SAFEARRAY::TYPE_POINTERS);
    return static_cast<unsigned int>(m_ptr->pointers.size());
}
