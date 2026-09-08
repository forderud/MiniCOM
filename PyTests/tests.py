import PyTests # require PyTests.pyd in PYTHONPATH or sys.path

calc = PyTests.CreateCalculator()

val = calc.GetValue()
print("calc.GetValue() returned "+str(val))

val = calc.Add(1,2)
print("calc.Add(1, 2) returned "+str(val))
