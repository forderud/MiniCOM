import PyTests # require PyTests.pyd in PYTHONPATH or sys.path

calc = PyTests.CreateCalculator()

val = calc.GetValue()
print("calc.GetValue() returned "+str(val))

val = calc.Add(1,2)
print("calc.Add(1, 2) returned "+str(val))

calc2 = calc.QueryInterface(PyTests.ICalc2)
val = calc2.GetValue2()
print("calc2.GetValue() returned "+str(val))
