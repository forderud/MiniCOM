#include "NonWindows.hpp"


std::map<std::pair<GUID,std::string>, IUnknownFactory::Factory> & IUnknownFactory::Factories () {
    static std::map<std::pair<GUID,std::string>, Factory> s_factory;
    return s_factory;
}


template <> ATL::CComSafeArray<BSTR>::CComSafeArray (UINT size) {
    m_ptr.reset(new SAFEARRAY(SAFEARRAY::TYPE_STRINGS));
    m_ptr->strings.resize(size);
}
template <> ATL::CComSafeArray<IUnknown*>::CComSafeArray (UINT size) {
    m_ptr.reset(new SAFEARRAY(SAFEARRAY::TYPE_POINTERS));
    m_ptr->pointers.resize(size);
}

template <> ATL::CComTypeWrapper<BSTR>::type& CComSafeArray<BSTR>::GetAt (int idx) const {
    assert(m_ptr);
    assert(m_ptr->type == SAFEARRAY::TYPE_STRINGS);
    return m_ptr->strings[idx];
}
template <> ATL::CComTypeWrapper<IUnknown*>::type& CComSafeArray<IUnknown*>::GetAt (int idx) const {
    assert(m_ptr);
    assert(m_ptr->type == SAFEARRAY::TYPE_POINTERS);
    return m_ptr->pointers[idx];
}

template <> HRESULT ATL::CComSafeArray<BSTR>::Add (const typename CComTypeWrapper<BSTR>::type& t, BOOL copy) {
    (void)copy; // mute unreferenced argument warning
    assert(m_ptr);
    assert(m_ptr->type == SAFEARRAY::TYPE_STRINGS);
    m_ptr->strings.push_back(t);
    return S_OK;
}
template <> HRESULT ATL::CComSafeArray<IUnknown*>::Add (const typename CComTypeWrapper<IUnknown*>::type& t, BOOL copy) {
    (void)copy; // mute unreferenced argument warning
    assert(m_ptr);
    assert(m_ptr->type == SAFEARRAY::TYPE_POINTERS);
    m_ptr->pointers.push_back(t);
    return S_OK;
}

template <> unsigned int ATL::CComSafeArray<BSTR>::GetCount () const {
    assert(m_ptr);
    assert(m_ptr->type == SAFEARRAY::TYPE_STRINGS);
    return static_cast<unsigned int>(m_ptr->strings.size());
}
template <> unsigned int ATL::CComSafeArray<IUnknown*>::GetCount () const {
    assert(m_ptr);
    assert(m_ptr->type == SAFEARRAY::TYPE_POINTERS);
    return static_cast<unsigned int>(m_ptr->pointers.size());
}
