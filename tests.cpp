#include <cassert>
#include "NonWindows.hpp"


void TestCComSafeArray() {
    // create numeric SAFEARRAY object with 3 elements
    CComSafeArray<double> obj(3);
    {
        double * vals = &obj.GetAt(0); // will fail if empty
        vals[0] = 2.0;
        vals[1] = 3.0;
        vals[2] = 4.0;
    }
    
    assert(obj.GetCount() == 3u);
}

int main() {
    printf("Running tests...\n");
    TestCComSafeArray();
}
