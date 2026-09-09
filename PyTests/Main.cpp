#include <AppAPI/ComSupport.hpp>
#include "Example.tlh"
#include <pybind11/pybind11.h>
#include <memory>

namespace py = pybind11;
using namespace TestInterfaces;

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

    HRESULT raw_GetValue (/*out*/int * value) override {
        if (!value)
            return E_INVALIDARG;

        *value = 42;
        return S_OK;
    }

    HRESULT raw_Add (int a, int b, /*out*/int * result) override {
        if (!result)
            return E_INVALIDARG;

        *result = a + b;
        return S_OK;
    }

    HRESULT raw_GetValue2 (/*out*/int * value) override {
        if (!value)
            return E_INVALIDARG;

        *value = 43;

        if (m_callback)
            m_callback->Message(_bstr_t(L"GetValue2 called"));

        return S_OK;
    }

    HRESULT raw_SetCallback(ICalcCb* cb) override {
        if (!cb)
            return E_INVALIDARG;

        m_callback = cb;
        return S_OK;
    }

    BEGIN_COM_MAP(Calculator)
        COM_INTERFACE_ENTRY(ICalc)
        COM_INTERFACE_ENTRY(ICalcExt)
        COM_INTERFACE_ENTRY(ICalc2)
    END_COM_MAP()

private:
    CComPtr<ICalcCb> m_callback;
};


/** Trampoline class for Python callbacks. */
class PyICalcCb :
    public CComObjectRootEx<CComMultiThreadModel>, // also compatible with STA
    public ICalcCb {
public:
    PyICalcCb() {
        py::print("PyICalcCb ctor.\n");
    }
    ~PyICalcCb() {
        py::print("PyICalcCb dtor.\n");
    }

    // Trampoline for the pure virtual function
    HRESULT raw_Message(BSTR msg) override {
        PYBIND11_OVERRIDE_PURE(
            HRESULT,  // Return type
            ICalcCb,  // Parent class
            Message,  // Name of function in C++ (and Python)
            msg       // Arguments
        );
    }

    BEGIN_COM_MAP(PyICalcCb)
        COM_INTERFACE_ENTRY(ICalcCb)
    END_COM_MAP()
};

/** Declare CComPtr<T> as smart-pointer type. */
PYBIND11_DECLARE_HOLDER_TYPE(T, CComPtr<T>, /*intrusive ref-count*/true);

namespace PYBIND11_NAMESPACE {
    namespace detail {
        /** Work-around for missing a .get() method in CComPtr<T>. */
        template <typename T>
        struct holder_helper<CComPtr<T>> { // <-- specialization
            static T* get(const CComPtr<T>& p) {
                return p;
            }
        };
    }
}

template <class T>
static CComPtr<T> ComCast(IUnknown& obj) {
    T* raw = nullptr;
    CHECK(obj.QueryInterface(&raw));

    CComPtr<T> ptr;
    ptr.Attach(raw); // take ownership without an extra AddRef
    return ptr;
}

/** Map a python interface type to a QueryInterface call. */
static py::object QueryInterface(IUnknown& obj, const py::object& iface) {
    if (iface.is(py::type::of<IUnknown>()))
        return py::cast(ComCast<IUnknown>(obj));
    if (iface.is(py::type::of<ICalcCb>()))
        return py::cast(ComCast<ICalcCb>(obj));
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
    py::class_<IUnknown, CComPtr<IUnknown>>(m, "IUnknown")
        .def("QueryInterface", &QueryInterface, "Type cast method")
        .def("AddRef", &IUnknown::AddRef)
        .def("Release", &IUnknown::Release);

    /** Bind ICalcCb. */
    py::class_<ICalcCb, CComPtr<ICalcCb>>(m, "ICalcCb")
        .def(py::init([]() {
            // CComObject starts at zero references; pybind11's holder then takes one
            CComObject<PyICalcCb>* obj = nullptr;
            CHECK(CComObject<PyICalcCb>::CreateInstance(&obj));
            return static_cast<ICalcCb*>(obj);
            }))
        .def("Message", &ICalcCb::Message);

    /** Bind ICalc. */
    py::class_<ICalc, IUnknown, CComPtr<ICalc>>(m, "ICalc")
        .def("GetValue", &ICalc::GetValue);

    /** Bind ICalcExt. */
    py::class_<ICalcExt, ICalc, CComPtr<ICalcExt>>(m, "ICalcExt")
        .def("Add", &ICalcExt::Add)
        .def("SetCallback", &ICalcExt::SetCallback);

    /** Bind ICalc2.. */
    py::class_<ICalc2, IUnknown, CComPtr<ICalc2>>(m, "ICalc2")
        .def("GetValue2", &ICalc2::GetValue2);

    /** Factory function. */
    m.def("CreateCalculator", []() {
        CComPtr<Calculator> calculator = CreateLocalInstance<Calculator>();
        CComPtr<ICalcExt> ptr;
        CHECK(calculator->QueryInterface(__uuidof(ICalcExt), (void**)&ptr));
        return ptr;
    }, "Create COM Calculator object");
}
