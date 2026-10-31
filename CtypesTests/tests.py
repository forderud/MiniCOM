# Calls a COM object from Python, through bindings generated from the IDL.
#
# Does the same as PyTests/tests.py, which needs a pybind11 module written by
# hand for these interfaces. Nothing here is written for them: ExampleBindings
# comes out of the Example namespace of TestInterfaces.winmd, which the Windows
# build describes Example.idl in, and the object is reached through ctypes.
#
# What the .winmd does not keep is checked too, as what a caller gets instead,
# so that it shows up here if a later generator starts keeping it.

import ctypes
import gc
import sys

import comruntime
import ExampleBindings # require ExampleBindings.py in PYTHONPATH or sys.path
import winmd2py


def LeftOut (call, reason):
    '''Checks that a call needs what the .winmd does not keep.'''
    try:
        call()
        assert False, "should not be callable"
    except NotImplementedError as error:
        assert reason in str(error), str(error)


class CalcCb:
    def Message(self, msg):
        print("Received message: "+msg)


cb = ExampleBindings.ImplementICalcCb(CalcCb())

library = ctypes.CDLL(sys.argv[1])
calc = ExampleBindings.CreateCalculator(library, ExampleBindings.ICalcExt)
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
# the imported file, IUnknown, strings, arrays, SAFEARRAYs, booleans, enums and
# structs holding all of those
import RegistryBindings
from RegistryBindings import Record, Span, Shape, SHAPE_LINE, SHAPE_BOX

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

# SAFEARRAYs. MIDL writes SAFEARRAY(BYTE) into the header as "SAFEARRAY *", so
# the element type that the bindings need to pass one never reaches the .winmd
LeftOut(lambda: registry.Reverse(b"\x01\x02\x03"), "SAFEARRAY")
LeftOut(lambda: registry.Upper(["ab", "Cd"]), "SAFEARRAY")
LeftOut(lambda: registry.Objects(), "SAFEARRAY")

# a struct with every kind of field, made in C++ and read back by it. It holds
# a SAFEARRAY, so it is left out as a whole, and every method that passes it
LeftOut(lambda: registry.MakeRecord("made in C++"), "SAFEARRAY")
LeftOut(lambda: registry.CheckRecord(Record()), "SAFEARRAY")

# without its SAFEARRAY it would still come out wrong: BOOL and long are "long"
# in C, 64bit on 64bit Linux and macOS. BOOL is a typedef that keeps its name in
# the .winmd, but a plain long is only recorded as Int32, its size on Windows.
fields = {}
for described in winmd2py.Winmd(sys.argv[2]).Types("Registry"):
    if described["Name"] == "Record":
        for field in described["Fields"]:
            fields[field["Name"]] = field["Type"]
assert fields["flag"] == {"Kind": "Native", "Name": "BOOL"}
assert fields["big"] == {"Kind": "Native", "Name": "int"}, fields["big"] # declared "long"
assert fields["weights"] == {"Kind": "Array", "Shape": {"Size": 3}, "Child": {"Kind": "Native", "Name": "float"}}

stats = registry.QueryInterface(RegistryBindings.IStats)

# a class can also be created by name, the way a library whose IDL declares no
# coclass is used
for progid in ["RegistryImpl", "Example.RegistryImpl.1"]:
    other = comruntime.CreateByName(library, progid, RegistryBindings.IStats)
    assert other.StoreCount() == 0
    del other
val = stats.Sum([1.0, 2.0, 3.5])
print("stats.Sum([1, 2, 3.5]) returned "+str(val))
assert val == 6.5
assert stats.StoreCount() == 1

# several [out] arguments come back as a tuple, an [out] array as a list, and
# an [in,out] argument is passed in and returned
assert stats.MinMax([3.0, -1.0, 2.0]) == (-1.0, 3.0)
assert stats.Bytes(0x04030201) == [1, 2, 3, 4]
assert stats.Scale(2.5, True) == -5.0
assert stats.Scale(2.5, False) == 5.0
assert stats.Positive(1.0) is True
assert stats.Positive(-1.0) is False

# structs by value and returned, and enums, where a value that is no member,
# such as a combination of flags, comes back as a plain int
assert stats.GetSpan([3.0, -1.0, 2.0]) == Span(low=-1.0, high=3.0)
assert stats.Widen(Span(low=1.0, high=2.0), 0.5) == Span(low=0.5, high=2.5)
assert stats.Next(SHAPE_LINE) is Shape.SHAPE_BOX
assert stats.Next(SHAPE_BOX) == 4

# a method with an argument the bindings cannot pass says so, and keeps its
# place in the vtable, so the methods after it still reach the right slot
LeftOut(lambda: stats.Opaque(None), "int **")

# a Python object kept by C++ as IUnknown: the registry holds a reference while
# it keeps it, and fetching it back hands out the same object
cb2 = ExampleBindings.ImplementICalcCb(CalcCb())
registry.Store(cb2)
assert cb2.references == 2
fetched = registry.Fetch()
assert fetched.pointer.value == cb2.pointer.value
del fetched
assert cb2.references == 2
assert stats.StoreCount() == 2


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

    def Samples(self):
        return [0.5, 1.5, 2.0]

    def Label(self, words):
        lengths = []
        for word in words:
            lengths.append(float(len(word)))
        return Record(name=" ".join(words), shape=SHAPE_LINE, span=Span(low=0.0, high=len(words)),
                      samples=lengths, valid=True, flag=1, big=1 << 40, after=2.5)

# C++ calls each method in turn, and stops at Samples, which hands out a
# SAFEARRAY and so answers E_NOTIMPL from Python
calc2 = registry.NewCalculator().QueryInterface(ExampleBindings.ICalc2)
source = RegistryBindings.ImplementISource(Source(calc2))
try:
    registry.Describe(source)
    assert False, "Describe should have stopped at Samples"
except comruntime.COMError as error:
    assert error.hr == 0x80004001, str(error) # E_NOTIMPL
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

del registry, stats
gc.collect()
assert len(comruntime._held) == 0 # and is let go of with the registry
assert cb2.references == 1
