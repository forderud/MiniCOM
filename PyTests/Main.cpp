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
        printf("Calculator ctor.\n");
    }

    ~Calculator() {
        printf("Calculator dtor.\n");
    }

    HRESULT GetValue (/*out*/int * value) override {
        *value = 42;
        return S_OK;
    }

    HRESULT Add (int a, int b, /*out*/int * result) override {
        *result = a + b;
        return S_OK;
    }

    HRESULT GetValue2 (/*out*/int * value) override {
        *value = 43;
        return S_OK;
    }

    BEGIN_COM_MAP(Calculator)
        COM_INTERFACE_ENTRY(ICalc)
        COM_INTERFACE_ENTRY(ICalcExt)
        COM_INTERFACE_ENTRY(ICalc2)
    END_COM_MAP()
};


/** Declare CComPtr<T> as smart-pointer type. */
PYBIND11_DECLARE_HOLDER_TYPE(T, CComPtr<T>, /*intrusive ref-count*/true);

namespace PYBIND11_NAMESPACE {
    namespace detail {
        /** Work-around for missing a .get() method in CComPtr<T>. */
        template <typename T>
        struct holder_helper<CComPtr<T>> { // <-- specialization
            static const T* get(const CComPtr<T>& p) {
                return p;
            }
        };
    }
}

PYBIND11_MODULE(PyTests, m, py::mod_gil_not_used()) {
    /** Bind IUnknown. */
    py::class_<IUnknown, CComPtr<IUnknown>>(m, "IUnknown")
        .def("AddRef", &IUnknown::AddRef)
        .def("Release", &IUnknown::Release);

    /** Bind ICalcExt. */
    py::class_<ICalcExt, CComPtr<ICalcExt>>(m, "ICalcExt")
        .def("GetValue", [](ICalcExt* m) {
            // convert output argument to return value
            int val = 0;
            CHECK(m->GetValue(&val));
            return val;
        })
        .def("Add", [](ICalcExt* m, int left, int right) {
            // convert output argument to return value
            int val = 0;
            CHECK(m->Add(left, right, &val));
            return val;
        });

    /** Factory function. */
    m.def("CreateCalculator", []() {
        CComPtr<Calculator> calculator = CreateLocalInstance<Calculator>();
        CComPtr<ICalcExt> ptr;
        CHECK(calculator.QueryInterface(&ptr));
        return ptr;
        }, "Create COM Calculator object");
}
