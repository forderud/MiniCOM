#pragma once
#include <atomic>
#ifdef _WIN32
  #include <atlcom.h>
#else
  #include "NonWindows.hpp"
#endif


/** IUnknown-based alternative to Microsoft's IWeakReference.
    References to this interface are weak, meaning they do not extend the lifetime of the object.
    Casting back to IUnknown and other interfaces only succeed if the object is still alive. */
struct DECLSPEC_UUID("146532F9-763D-44C9-875A-7B5B732B9046")
IWeakRef : public IUnknown {
};

/** COM wrapper class that provides support for weak references through the IWeakRef interface. */
class SharedRefBase : public IUnknown {
public:
    SharedRefBase() : m_weak(*this) {
        ++s_obj_count;
    }
    virtual ~SharedRefBase() {
        --s_obj_count;
    }

    /** QueryInterface doesn't adhere to the COM rules, since the interfaces accessible are dynamic.
        Still, interface disappearance only affect IWeakRef clients and is not noticable for clients unaware of this interface.
        DOC: https://learn.microsoft.com/en-us/windows/win32/com/rules-for-implementing-queryinterface */
    HRESULT QueryInterface(const IID& iid, void** ptr) override {
        if (!ptr)
            return E_INVALIDARG;

        *ptr = nullptr;
        if (iid == IID_IUnknown) {
            if (!Inner())
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
        } else {
            if (!Inner())
                return E_NOT_SET;

            return Inner()->QueryInterface(iid, ptr);
        }
    }

    ULONG AddRef() override {
        return m_refs.AddRef(true);
    }

    ULONG Release() override {
        ULONG strong = m_refs.Release(true);

        if (!strong) {
            if (Inner()) {
                // release object handle
                auto* ptr = Inner();
                ClearInner();
                ptr->Release(); // might trigger reentrancy, so don't touch members afterwards
            }

            m_weak.Release(); // also release weak reference when strong=0
        }

        return strong;
    }

    static ULONG ObjectCount() {
        return s_obj_count;
    }

protected:
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
        HRESULT QueryInterface(const IID& iid, void** ptr) override {
            if (!ptr)
                return E_INVALIDARG;

            *ptr = nullptr;
            if (iid == __uuidof(IWeakRef)) {
                // special handling of weak reference to object
                *ptr = static_cast<IWeakRef*>(this);
                m_parent.m_refs.AddRef(false);
                return S_OK;
            } else {
                // all other interfaces require a preexisting strong reference
                if (!m_parent.Inner())
                    return E_NOT_SET;

                // add temporarily strong reference to avoid concurrent deletion by other threads 
                ULONG strong = m_parent.AddRef();
                if (strong <= 1) {
                    // object was just destroyed
                    m_parent.Release();
                    return E_NOT_SET;
                }

                // forward call to parent object
                HRESULT hr = m_parent.QueryInterface(iid, ptr);

                // release temporary strong reference
                m_parent.Release();

                return hr;
            }
        }

        ULONG AddRef() override {
            return m_parent.m_refs.AddRef(false);
        }

        ULONG Release() override {
            ULONG weak = m_parent.m_refs.Release(false);
            if (!weak)
                delete &m_parent;

            return weak;
        }

    private:
        SharedRefBase& m_parent;
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

        ULONG Release(bool _strong) {
            if (_strong) {
                //std::cout << "  strong=" << (strong - 1) << " weak=" << weak << std::endl;
                return --strong;
            } else {
                //std::cout << "  strong=" << strong << " weak=" << (weak - 1) << std::endl;
                return --weak;
            }
        }

    protected:
        std::atomic<uint32_t> strong = 0; // strong ref-count for m_inner lifetime
        std::atomic<uint32_t> weak = 1;   // weak ref-count for SharedRefBase lifetime + 1 if strong>0
    };

    /** Get or set pointer to inner object. */
    virtual IUnknown* Inner() = 0;
    virtual void     ClearInner() = 0;

    AtomicRefBlock        m_refs; // Reference-counts. Only touched once per method for thread safety.
    WeakRef               m_weak;
    static inline ULONG   s_obj_count = 0;
};


/** Create a weak-pointer compatible aggregated COM object from a C++ COM class. */
template <class Class>
class SharedRef : public SharedRefBase {
public:
    SharedRef() {
        CComAggObject<Class>::CreateInstance(this, &m_inner);
        assert(m_inner);
        m_inner->AddRef(); // doesn't increase ref-count for this
    }
    ~SharedRef() override {
        assert(m_inner == nullptr);
    }

    /** Pointer to internal C++ class. Potentially unsafe. Only call immediately after construction. */
    Class* Internal() const {
        assert(m_inner);
        return &m_inner->m_contained;
    }

private:
    IUnknown* Inner() override {
        return m_inner;
    }
    void ClearInner() override {
        m_inner = nullptr;
    }

    CComAggObject<Class>* m_inner = nullptr;
};


/** Create a weak-pointer compatible aggregated COM object based on ClassID. */
class SharedRefClsid : public SharedRefBase {
public:
    SharedRefClsid(GUID clsid, DWORD context = CLSCTX_ALL) {
        CComPtr<IUnknown> inner;
        CHECK(inner.CoCreateInstance(clsid, this, context));
        m_inner = inner.Detach();
        assert(m_inner);
    }
    ~SharedRefClsid() override {
        assert(m_inner == nullptr);
    }

private:
    IUnknown* Inner() override {
        return m_inner;
    }
    void ClearInner() override {
        m_inner = nullptr;
    }

    IUnknown* m_inner = nullptr;
};
