#pragma once
#include <atomic>
#include <iostream>
#include "ComSupport.hpp"


/** IUnknown-based alternative to Microsoft's IWeakReference.
    References to this interface are weak, meaning they do not extend the lifetime of the object.
    Casting back to IUnknown and other interfaces only succeed if the object is still alive. */
struct DECLSPEC_UUID("146532F9-763D-44C9-875A-7B5B732B9046")
IWeakRef : public IUnknown {
public:
};


/** Base-class for COM wrapper that provides support for weak references through the IWeakRef interface. */
template <class Class>
class SharedRefBase : public IUnknown {
public:
    SharedRefBase() : m_weak(*this) {
        CComAggObject<Class>::CreateInstance(this, &m_ptr);
        assert(m_ptr);
        m_ptr->AddRef(); // doesn't increase ref-count for this
        ++s_obj_count;
    }
    virtual ~SharedRefBase() {
        assert(m_ptr == nullptr);
        --s_obj_count;
    }

    /** QueryInterface doesn't adhere to the COM rules, since the interfaces accessible are dynamic.
        Still, interface disappearance only affect IWeakRef clients and is not noticable for clients unaware of this interface.
        DOC: https://learn.microsoft.com/en-us/windows/win32/com/rules-for-implementing-queryinterface */
    HRESULT QueryInterfaceInternal(const IID& iid, void** ptr) {
        if (!ptr)
            return E_INVALIDARG;

        *ptr = nullptr;
        if (iid == IID_IUnknown) {
            if (!m_ptr)
                return E_NOT_SET;

            // strong reference to this
            *ptr = static_cast<IUnknown*>(this);
            m_refs.AddRef(true);
            return S_OK;
        } else if (iid == __uuidof(IWeakRef)) {
            // weak reference to child object
            *ptr = static_cast<IWeakRef*>(&m_weak);
            m_refs.AddRef(false);
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    ULONG AddRef() override {
        return m_refs.AddRef(true);
    }

    ULONG Release() override {
        RefBlock refs = m_refs.Release(true);

        if (m_ptr && !refs.strong) {
            // release object handle
            m_ptr->Release();
            m_ptr = nullptr;
        }

        if (!refs.strong && !refs.weak)
            delete this;

        return refs.strong;
    }

    static ULONG ObjectCount() {
        return s_obj_count;
    }

protected:
    template <class T>
    HRESULT CastAggObject(/*out*/void** ptr) {
        *ptr = nullptr;
        if (!m_ptr)
            return E_NOT_SET;

        // return strong reference pointed-to object
        *ptr = static_cast<T*>(&m_ptr->m_contained);
        m_refs.AddRef(true);
        return S_OK;
    }

private:
    /** Inner class for managing weak references. */
    class WeakRef : public IWeakRef {
        friend class SharedRefBase;
    public:
        WeakRef(SharedRefBase& parent) : m_parent(parent) {
        }
        ~WeakRef() {
        }

        /** QueryInterface doesn't adhere to the COM aggregation rules for inner objects, since it lacks special handling of IUnknown.
            Still, this only affect IWeakRef clients and is not noticable for clients unaware of this interface.
            DOC: https://learn.microsoft.com/en-us/windows/win32/com/aggregation */
        HRESULT QueryInterface(const IID& iid, void** obj) override {
            return m_parent.QueryInterface(iid, obj);
        }

        ULONG AddRef() override {
            return m_parent.m_refs.AddRef(false);
        }

        ULONG Release() override {
            RefBlock refs = m_parent.m_refs.Release(false);
            if (!refs.strong && !refs.weak)
                delete &m_parent;

            return refs.weak;
        }

    private:
        SharedRefBase& m_parent;
    };

    struct RefBlock {
        uint32_t strong = 0; // strong use-count for m_ptr lifetime
        uint32_t weak = 0;   // weak ref-count for SharedRef lifetime
    };
    /** Thread-safe handling of strong & weak reference-counts. */
    struct AtomicRefBlock {
    public:
        ULONG AddRef(bool _strong) {
            if (_strong) {
                //std::cout << "  strong=" << (strong + 1) << " weak=" << weak << std::endl;
                return ++strong;
            } else {
                //std::cout << "  strong=" << strong << " weak=" << (weak + 1) << std::endl;
                return ++weak;
            }
        }

        /** Returns a copy of the struct to facillitate a thread safe parent class. */
        RefBlock Release(bool _strong) {
            if (_strong) {
                //std::cout << "  strong=" << (strong - 1) << " weak=" << weak << std::endl;
                return { --strong, weak };
            } else {
                //std::cout << "  strong=" << strong << " weak=" << (weak - 1) << std::endl;
                return { strong, --weak };
            }
        }

    private:
        std::atomic<uint32_t> strong = 0; // strong ref-count for m_ptr lifetime
        std::atomic<uint32_t> weak = 0;   // weak ref-count for SharedRef lifetime
    };

    AtomicRefBlock        m_refs; // Reference-counts. Only touched once per method for thread safety.
    CComAggObject<Class>* m_ptr = nullptr;
    WeakRef               m_weak;
    static inline ULONG   s_obj_count = 0;
};


/** COM wrapper class that provides support for weak references through the IWeakRef interface. */
template <class Class, class Interface1>
class SharedRef1 : public SharedRefBase<Class> {
    using BASE = SharedRefBase<Class>;
public:
    HRESULT QueryInterface(const IID& iid, void** ptr) override {
        if (iid == __uuidof(Interface1)) {
            return BASE::template CastAggObject<Interface1>(ptr);
        } else {
            return BASE::QueryInterfaceInternal(iid, ptr);
        }
    }
};

/** COM wrapper class that provides support for weak references through the IWeakRef interface. */
template <class Class, class Interface1, class Interface2>
class SharedRef2 : public SharedRefBase<Class> {
    using BASE = SharedRefBase<Class>;
public:
    HRESULT QueryInterface(const IID& iid, void** ptr) override {
        if (iid == __uuidof(Interface1)) {
            return BASE::template CastAggObject<Interface1>(ptr);
        } else if (iid == __uuidof(Interface2)) {
            return BASE::template CastAggObject<Interface2>(ptr);
        } else {
            return BASE::QueryInterfaceInternal(iid, ptr);
        }
    }
};

/** COM wrapper class that provides support for weak references through the IWeakRef interface. */
template <class Class, class Interface1, class Interface2, class Interface3>
class SharedRef3 : public SharedRefBase<Class> {
    using BASE = SharedRefBase<Class>;
public:
    HRESULT QueryInterface(const IID& iid, void** ptr) override {
        if (iid == __uuidof(Interface1)) {
            return BASE::template CastAggObject<Interface1>(ptr);
        } else if (iid == __uuidof(Interface2)) {
            return BASE::template CastAggObject<Interface2>(ptr);
        } else if (iid == __uuidof(Interface3)) {
            return BASE::template CastAggObject<Interface3>(ptr);
        } else {
            return BASE::QueryInterfaceInternal(iid, ptr);
        }
    }
};
