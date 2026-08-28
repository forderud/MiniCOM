#include <cassert>
#include <vector>
#include "NonWindows.hpp"


/** Convert raw array to SafeArray. */
template <class T>
CComSafeArray<T> ConvertToSafeArray (const T * input, size_t element_count) {
    CComSafeArray<T> result(static_cast<ULONG>(element_count));

    if (element_count > 0) {
        T * out_ptr = &result.GetAt(0); // will fail if empty
        for (size_t i = 0; i < element_count; ++i)
            out_ptr[i] = input[i];
    }
    return result;
}


void TestCComSafeArray() {
    std::vector<double> vals = {2.0, 3.0, 4.0};
    {
        printf("direct assigmnent...\n");
        CComSafeArray<double> sa_vals = ConvertToSafeArray(vals.data(), vals.size());
        assert(sa_vals.GetCount() == 3);
    }
    {
        printf("copy-assignment...\n");
        CComSafeArray<double> sa_vals;
        sa_vals = ConvertToSafeArray(vals.data(), vals.size());
        assert(sa_vals.GetCount() == 3);
    }
}

int main() {
    printf("Running tests...\n");
    TestCComSafeArray();
}
