/** Example of cross-platform Python bindings.
*   Based on https://pybind11.readthedocs.io/en/stable/advanced/classes.html */

#include <AppAPI/ComSupport.hpp>
#include "Example.h"
#include <gtest/gtest.h>
#include <pybind11/embed.h>

namespace py = pybind11;


/** Trampoline class to redirect virtual calls to Python. */
class CalcExtTrampoline : public ICalcExt, public py::trampoline_self_life_support {
public:
    /* IUnknown base interface */
    HRESULT QueryInterface(const GUID& iid, /*out*/void** obj) override {
        PYBIND11_OVERRIDE_PURE(
            HRESULT,  // Return type
            ICalcExt, // Parent class
            QueryInterface, // Function name
            iid, obj  // Argument(s)
        );
    }
    ULONG AddRef() override {
        PYBIND11_OVERRIDE_PURE(
            HRESULT,  // Return type
            ICalcExt, // Parent class
            AddRef, // Function name
            // Argument(s)
        );
    }
    ULONG Release() override {
        PYBIND11_OVERRIDE_PURE(
            HRESULT,  // Return type
            ICalcExt, // Parent class
            Release,  // Function name
            // Argument(s)
        );
    }

    /* ICalcExt interface */
    HRESULT GetValue(/*out*/int* result) override {
        PYBIND11_OVERRIDE_PURE(
            HRESULT,  // Return type
            ICalcExt, // Parent class
            GetValue, // Function name
            result    // Argument(s)
        );
    }
    HRESULT Add(int left, int right, /*out*/int* result) override {
        PYBIND11_OVERRIDE_PURE(
            HRESULT,  // Return type
            ICalcExt, // Parent class
            Add,      // Function name
            left, right, result // Argument(s)
        );
    }
};

PYBIND11_EMBEDDED_MODULE(example, m) {
    py::class_<ICalcExt, CalcExtTrampoline, py::smart_holder>(m, "ICalcExt")
        .def(py::init<>())
        .def("GetValue", &ICalcExt::GetValue)
        .def("Add", &ICalcExt::Add);
}

#if 0
class PyCalcExt(ICalcExt):
    def QueryInterface(self, iid, obj):
        return 0x80004005L # E_FAIL
    def AddRef(self):
        return 1
    def Release(self):
        return 1
    def GetValue(self, result):
        result = 41
        return 0 #S_OK
    def Add(self, left, right, result):
        result = left + right
        return 0 #S_OK
#endif

TEST(PythonBindings, TestPythonInterpreter) {
    // 1. Start the Python interpreter
    py::scoped_interpreter guard{};

    try {
        // 2. Import a module or run raw string scripts
        py::exec("print('Hello from Python inside C++!')");

        // 3. Call a function from a local script (e.g., myscript.py)
        py::module_ sys = py::module_::import("sys");
        sys.attr("path").attr("append")("."); // Add current directory to path

        py::module_ module = py::module_::import("example");
        py::object calc = module.attr("ICalcExt");
        //py::object value = calc.call("GetValue()");
#if 0
        py::object result = my_module.attr("add_numbers")(5, 10);

        // 4. Cast the result back to C++ types
        std::cout << "Result from Python: " << result.cast<int>() << std::endl;
#endif
    }
    catch (py::error_already_set& e) {
        std::cerr << "Python Error: " << e.what() << std::endl;
    }
}
