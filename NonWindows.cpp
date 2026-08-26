#include "NonWindows.hpp"


__attribute__((visibility("default")))
Buffer<IUnknownFactory::Entry> & IUnknownFactory::Factories () {
    static Buffer<Entry> s_factory;
    return s_factory;
}


/* ole32-parity object creation. Defined here, beside Factories(), because this
   is the one translation unit compiled into the library -- the rest of MiniCOM
   is header-only and so exports nothing a C caller can link against. */
extern "C" __attribute__((visibility("default")))
HRESULT CoCreateInstance (const GUID& rclsid, IUnknown* pUnkOuter,
                          DWORD /*dwClsContext*/, const GUID& riid, void** ppv) {
    if (!ppv)
        return E_POINTER;
    *ppv = nullptr;

    // Look the class up before creating it. IUnknownFactory::CreateInstance
    // asserts on an unregistered CLSID, which is a reasonable debugging aid for
    // C++ callers but would abort a caller that is entitled to an HRESULT --
    // ole32 returns REGDB_E_CLASSNOTREG here.
    bool registered = false;
    for (size_t i = 0; i < IUnknownFactory::Factories().size(); i++) {
        if (IUnknownFactory::Factories()[i].clsid == rclsid) {
            registered = true;
            break;
        }
    }
    if (!registered)
        return REGDB_E_CLASSNOTREG;

    IUnknown* obj = IUnknownFactory::CreateInstance(rclsid, pUnkOuter);
    if (!obj)
        return E_FAIL;

    // Hand back the interface asked for, as ole32 does, rather than IUnknown.
    HRESULT hr = obj->QueryInterface(riid, ppv);
    obj->Release();                 // QueryInterface took its own reference
    return hr;
}

extern "C" __attribute__((visibility("default")))
HRESULT CLSIDFromProgID (const wchar_t* lpszProgID, GUID* lpclsid) {
    if (!lpszProgID || !lpclsid)
        return E_POINTER;

    for (size_t i = 0; i < IUnknownFactory::Factories().size(); i++) {
        const auto& elm = IUnknownFactory::Factories()[i];
        if (elm.name == ATL::CComBSTR(lpszProgID)) {
            *lpclsid = elm.clsid;
            return S_OK;
        }
    }
    return CO_E_CLASSSTRING;
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
