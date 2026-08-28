#include <AppAPI/ComSupport.hpp>
#include <gtest/gtest.h>


static void CheckComErrorType(HRESULT hr1) {
    _com_error err(hr1);

    HRESULT hr2 = err.Error();
    EXPECT_EQ(hr1, hr2);

    // don't check string content, since it's not identical for Windows vs. other platforms.
    // Not stored in a typed variable: ErrorMessage() is const wchar_t* here, but const TCHAR*
    // on Windows, which is const char* in a non-Unicode build (cf. CHECK in ComSupport.hpp).
    EXPECT_TRUE(err.ErrorMessage());
}

TEST(ComErrorTests, Test_com_error) {
    // test 3 common error types
    CheckComErrorType(E_INVALIDARG);
    CheckComErrorType(E_BOUNDS);
    CheckComErrorType(E_FAIL);
}
