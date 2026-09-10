# Calls a COM object from Python, through bindings generated from the IDL.
#
# Does the same as PyTests/tests.py, which needs a pybind11 module written by
# hand for these interfaces. Nothing here is written for them: ExampleBindings
# comes out of Example.idl, and the object is reached through ctypes.

import ctypes
import sys

import comruntime
import ExampleBindings # require ExampleBindings.py in PYTHONPATH or sys.path


class CalcCb:
    def Message(self, msg):
        print("Received message: "+msg)


cb = ExampleBindings.ImplementICalcCb(CalcCb())

library = ctypes.CDLL(sys.argv[1])
calc = ExampleBindings.CreateCalculator(library, ExampleBindings.ICalcExt)
calc.SetCallback(cb.pointer)

val = calc.GetValue()
print("calc.GetValue() returned "+str(val))
assert val == 42

val = calc.Add(1,2)
print("calc.Add(1, 2) returned "+str(val))
assert val == 3

calc2 = calc.QueryInterface(ExampleBindings.ICalc2)
val = calc2.GetValue2()
print("calc2.GetValue2() returned "+str(val))
assert val == 43

# the object holds a reference on the Python callback for as long as it is
# alive, and the bindings drop the object when the last one goes out of scope
assert cb.references == 2
del calc, calc2
assert cb.references == 1
