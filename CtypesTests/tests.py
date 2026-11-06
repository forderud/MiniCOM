# Calls a COM object from Python, through bindings generated from the IDL.
#
# Does the same as PyTests/tests.py, which needs a pybind11 module written by
# hand for these interfaces. Nothing here is written for them: ExampleBindings
# comes out of the Example namespace of TestInterfaces.winmd, which the Windows
# build describes Example.idl in, and the object is reached through ctypes.

import ctypes
import gc
import sys

import comruntime
import ExampleBindings # require ExampleBindings.py in PYTHONPATH or sys.path
import RegistryBindings
from RegistryBindings import Record, Span, Shape, SHAPE_LINE, SHAPE_BOX


class CalcCb:
    def Message(self, msg):
        print("Received message: "+msg)


cb = ExampleBindings.ImplementICalcCb(CalcCb())

library = ctypes.CDLL(sys.argv[1])
calc = RegistryBindings.CreateCalculator(library, ExampleBindings.ICalcExt)
calc.SetCallback(cb)

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


# Registry.idl imports Example.idl and passes what it does not, the same kinds
# of arguments a real interface library passes around: interfaces declared in
# the imported file, IUnknown, strings, arrays, booleans, enums and structs
# holding all of those
registry = RegistryBindings.CreateRegistry(library, RegistryBindings.IRegistry)

# a calculator handed out as an interface from Example.idl, with its class
# from the bindings for that file
calc = registry.NewCalculator()
assert type(calc) is ExampleBindings.ICalcExt
val = calc.Add(2, 3)
print("registry.NewCalculator().Add(2, 3) returned "+str(val))
assert val == 5

# kept as IUnknown and fetched back, then reached through QueryInterface
registry.Store(calc)
del calc
calc2 = registry.Fetch().QueryInterface(ExampleBindings.ICalc2)
val = calc2.GetValue2()
print("registry.Fetch() -> ICalc2.GetValue2() returned "+str(val))
assert val == 43
del calc2

# a struct with every kind of field, made in C++ and read back by it
made = registry.MakeRecord("made in C++")
assert made == Record(name="made in C++", shape=SHAPE_BOX, span=Span(low=1.0, high=2.0),
                      weights=[0.5, 0.25, 0.25], valid=True, flag=1, after=2.5), made
assert registry.CheckRecord(made) == "made in C++ 2 1-2 0.5,0.25,0.25 valid 1 2.5"

# several [out] arguments come back as a tuple, an [out] array as a list, and
# an [in,out] argument is passed in and returned
assert registry.MinMax([3.0, -1.0, 2.0]) == (-1.0, 3.0)
assert registry.Bytes(0x04030201) == [1, 2, 3, 4]
assert registry.Scale(2.5, True) == -5.0
assert registry.Scale(2.5, False) == 5.0

# structs by value and returned, and enums, where a value that is no member,
# such as a combination of flags, comes back as a plain int
assert registry.Widen(Span(low=1.0, high=2.0), 0.5) == Span(low=0.5, high=2.5)
assert registry.Next(SHAPE_LINE) is Shape.SHAPE_BOX
assert registry.Next(SHAPE_BOX) == 4

# a method with an argument the bindings cannot pass says so, and keeps its
# place in the vtable, so the methods after it still reach the right slot
try:
    registry.Opaque(None)
    assert False, "Opaque should not be callable"
except NotImplementedError as error:
    assert "int **" in str(error), str(error)

# a Python object kept by C++ as IUnknown: the registry holds a reference while
# it keeps it, and fetching it back hands out the same object
cb2 = ExampleBindings.ImplementICalcCb(CalcCb())
registry.Store(cb2)
assert cb2.references == 2
fetched = registry.Fetch()
assert fetched.pointer.value == cb2.pointer.value
del fetched
assert cb2.references == 2


# a Python object that C++ calls into: ISource derives from ICalcCb, declared
# in Example.idl, and gets and hands out every kind of argument above
class Source:
    def __init__(self, calc):
        self.calc = calc
        self.messages = []
        self.accepted = None

    def Message(self, msg):
        self.messages.append(msg)

    def Name(self):
        return "python source"

    def Calculator(self):
        return self.calc

    def Accept(self, calc):
        self.accepted = calc # an [in] interface arrives as its class, and may be kept
        return calc.GetValue2() == 43

    def Range(self):
        return 1, 5

    def Adjust(self, value):
        return value * 2

    def Normalize(self, values):
        total = values[0] + values[1] + values[2]
        result = []
        for value in values:
            result.append(value / total)
        return result

    def GetSpan(self):
        return Span(low=0.5, high=4.0)

    def Label(self, words):
        return Record(name=words.upper(), shape=SHAPE_LINE, span=Span(low=0.0, high=len(words.split())),
                      weights=[0.5, 0.25, 0.25], valid=True, flag=1, after=2.5)

# C++ calls each method in turn, and reports what each returned
calc2 = registry.NewCalculator().QueryInterface(ExampleBindings.ICalc2)
source = RegistryBindings.ImplementISource(Source(calc2))
text = registry.Describe(source)
print("registry.Describe(source) returned "+text)
assert text == "python source uses 43, range 1-5, adjusted 3, normalized 0.125 0.25 0.625, span 0.5-4, label TWO WORDS 2 0.5", text
assert source.handler.messages == ["describing"]
assert source.references == 1
assert type(source.handler.accepted) is ExampleBindings.ICalc2
assert source.handler.accepted.GetValue2() == 43 # kept past the call that passed it
assert calc2.GetValue2() == 43 # the calculator handed to C++ was not released too often

# a handler that raises, or lacks the method, fails the call from C++ instead
# of looking like a success to it
class Failing(Source):
    def Name(self):
        raise ValueError("expected failure, for the test")

class Incomplete:
    def Message(self, msg):
        pass

for handler, hr in [(Failing(calc2), 0x80004005), (Incomplete(), 0x80004001)]:
    try:
        registry.Describe(RegistryBindings.ImplementISource(handler))
        assert False, "Describe should have failed"
    except comruntime.COMError as error:
        assert error.hr == hr, str(error)

# an object kept only by C++ stays alive, although Python has let go of it
registry.Store(RegistryBindings.ImplementISource(Source(calc2)))
gc.collect()
kept = registry.Fetch().QueryInterface(RegistryBindings.ISource)
assert kept.Name() == "python source"
del kept

del registry
gc.collect()
assert len(comruntime._held) == 0 # and is let go of with the registry
assert cb2.references == 1
