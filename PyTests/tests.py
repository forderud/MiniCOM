import PyTests # require PyTests.pyd in PYTHONPATH or sys.path


class CalcCb(PyTests.ICalcCb):
    def __init__(self):
        super().__init__()
    def Message(self, msg):
        print("Received message: "+msg)
        return 0 # S_OK

cb = CalcCb()

calc = PyTests.CreateCalculator()
calc.SetCallback(cb)

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
