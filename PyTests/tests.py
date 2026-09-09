import PyTests # require PyTests.pyd in PYTHONPATH or sys.path

calc = PyTests.CreateCalculator()

val = calc.GetValue()
print("calc.GetValue() returned "+str(val))
assert val == 42

val = calc.Add(1,2)
print("calc.Add(1, 2) returned "+str(val))
assert val == 3

calc2 = calc.QueryInterface(PyTests.ICalc2)
val = calc2.GetValue2()
print("calc2.GetValue2() returned "+str(val))
assert val == 43
