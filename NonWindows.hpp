#pragma once
/* ATL emulation for non-Windows platforms.

   Split out from the COM basics in NonWindowsCore.hpp, which is the unknwn.h
   equivalent. The code below calls IUnknown through its C++ interface, so it
   is unavailable when CINTERFACE is defined, mirroring Windows where such a
   translation unit gets unknwn.h but never the ATL headers. */

#include "NonWindowsCore.hpp"

#ifndef CINTERFACE

// error handler required by generated wrapper API headers
inline void _com_issue_errorex(HRESULT hr, IUnknown*, const IID &) {
    throw _com_error(hr);
}

template <class BASE>
class CComObject : public BASE {
public:
    CComObject() {
        static_assert(sizeof(*this) == sizeof(BASE)); // ensure it's safe to delete a BASE pointer without leaking memory
    }

    /* NO destructor here, since it's not guaranteed to be called. */

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
    bool operator==(const _com_ptr_t<U>& other) {
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

#endif // CINTERFACE
