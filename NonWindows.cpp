#include "NonWindows.hpp"


__attribute__((visibility("default")))
std::map<std::pair<GUID,std::string>, IUnknownFactory::Factory> & IUnknownFactory::Factories () {
    static std::map<std::pair<GUID,std::string>, Factory> s_factory;
    return s_factory;
}


template <> __attribute__((visibility("default")))
ATL::CComSafeArray<BSTR>::CComSafeArray (UINT size) {
    m_ptr.reset(new SAFEARRAY(SAFEARRAY::TYPE_STRINGS));
    m_ptr->strings.resize(size);
}
template <> __attribute__((visibility("default")))
ATL::CComSafeArray<IUnknown*>::CComSafeArray (UINT size) {
    m_ptr.reset(new SAFEARRAY(SAFEARRAY::TYPE_POINTERS));
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
HRESULT ATL::CComSafeArray<BSTR>::Add (const typename CComTypeWrapper<BSTR>::type& t, BOOL copy) {
    (void)copy; // mute unreferenced argument warning

    if (!m_ptr)
        m_ptr.reset(new SAFEARRAY(SAFEARRAY::TYPE_STRINGS)); // lazy initialization

    assert(m_ptr->type == SAFEARRAY::TYPE_STRINGS);
    m_ptr->strings.push_back(t);
    return S_OK;
}
template <> __attribute__((visibility("default")))
HRESULT ATL::CComSafeArray<IUnknown*>::Add (const typename CComTypeWrapper<IUnknown*>::type& t, BOOL copy) {
    (void)copy; // mute unreferenced argument warning

    if (!m_ptr)
        m_ptr.reset(new SAFEARRAY(SAFEARRAY::TYPE_POINTERS)); // lazy initialization

    assert(m_ptr->type == SAFEARRAY::TYPE_POINTERS);
    m_ptr->pointers.push_back(t);
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
