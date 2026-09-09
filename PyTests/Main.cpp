#include <AppAPI/ComSupport.hpp>
#include "Example.h"
#include <pybind11/pybind11.h>
#include <memory>

namespace py = pybind11;

#ifdef _WIN32
class PyTestsAtlModule : public ATL::CAtlDllModuleT<PyTestsAtlModule> {
};
PyTestsAtlModule _AtlModule;
#endif


/** C++ COM class to be called from Python. */
class Calculator :
    public CComObjectRootEx<CComMultiThreadModel>, // also compatible with STA
    public CComCoClass<Calculator>, // no CLSID needed
    public ICalcExt, public ICalc2 {
public:
    Calculator() {
        py::print("Calculator ctor.\n");
    }

    ~Calculator() {
        py::print("Calculator dtor.\n");
    }

    HRESULT GetValue (/*out*/int * value) override {
        if (!value)
            return E_INVALIDARG;

        *value = 42;
        return S_OK;
    }

    HRESULT Add (int a, int b, /*out*/int * result) override {
        if (!result)
            return E_INVALIDARG;

        *result = a + b;
        return S_OK;
    }

    HRESULT GetValue2 (/*out*/int * value) override {
        if (!value)
            return E_INVALIDARG;

        *value = 43;
        return S_OK;
    }

    HRESULT SetCallback(ICalcCb* cb) override {
        return E_NOTIMPL;
    }

    BEGIN_COM_MAP(Calculator)
        COM_INTERFACE_ENTRY(ICalc)
        COM_INTERFACE_ENTRY(ICalcExt)
        COM_INTERFACE_ENTRY(ICalc2)
    END_COM_MAP()
};


/** PyBind11-compatible wrapper around CComPtr<T>. */
template <typename T>
class CComHolder : public CComPtr<T> {
public:
    using CComPtr<T>::CComPtr;

    CComHolder() = default;

    CComHolder(const CComPtr<T>& other) : CComPtr<T>(other) {}
    CComHolder(CComPtr<T>&& other) : CComPtr<T>(other) {
    }

    /** Work-around for CComPtr::operator& asserting p==NULL. */
    CComHolder* operator&() noexcept {
        return this;
    }

    /** Work-around for missing CComPtr<T>::get() method. */
    T* get() const {
        return *this;
    }
};

/** Declare CComHolder<T> as smart-pointer type. */
PYBIND11_DECLARE_HOLDER_TYPE(T, CComHolder<T>, /*intrusive ref-count*/true);

template <class T>
static CComHolder<T> ComCast(IUnknown& obj) {
    CComHolder<T> ptr;
    CHECK(obj.QueryInterface(__uuidof(T), (void**)&ptr));
    return ptr;
}

/** Map a python interface type to a QueryInterface call. */
static py::object QueryInterface(IUnknown& obj, const py::object& iface) {
    if (iface.is(py::type::of<IUnknown>()))
        return py::cast(ComCast<IUnknown>(obj));
    if (iface.is(py::type::of<ICalc>()))
        return py::cast(ComCast<ICalc>(obj));
    if (iface.is(py::type::of<ICalcExt>()))
        return py::cast(ComCast<ICalcExt>(obj));
    if (iface.is(py::type::of<ICalc2>()))
        return py::cast(ComCast<ICalc2>(obj));

    throw py::type_error("Unknown COM interface.");
}

PYBIND11_MODULE(PyTests, m, py::mod_gil_not_used()) {
    /** Bind IUnknown. */
    py::class_<IUnknown, CComHolder<IUnknown>>(m, "IUnknown")
        .def("QueryInterface", &QueryInterface, "Type cast method")
        .def("AddRef", &IUnknown::AddRef)
        .def("Release", &IUnknown::Release);

    /** Bind ICalc. */
    py::class_<ICalc, IUnknown, CComHolder<ICalc>>(m, "ICalc")
        .def("GetValue", [](ICalc& self) {
            // convert output argument to return value
            int val = 0;
            CHECK(self.GetValue(&val));
            return val;
        });

    /** Bind ICalcExt. */
    py::class_<ICalcExt, ICalc, CComHolder<ICalcExt>>(m, "ICalcExt")
        .def("Add", [](ICalcExt& self, int left, int right) {
            // convert output argument to return value
            int val = 0;
            CHECK(self.Add(left, right, &val));
            return val;
        });

    /** Bind ICalc2.. */
    py::class_<ICalc2, IUnknown, CComHolder<ICalc2>>(m, "ICalc2")
        .def("GetValue2", [](ICalc2& self) {
            // convert output argument to return value
            int val = 0;
            CHECK(self.GetValue2(&val));
            return val;
        });

    /** Factory function. */
    m.def("CreateCalculator", []() {
        CComPtr<Calculator> calculator = CreateLocalInstance<Calculator>();
        CComHolder<ICalcExt> ptr;
        CHECK(calculator->QueryInterface(__uuidof(ICalcExt), (void**)&ptr));
        return ptr;
    }, "Create COM Calculator object");
}
