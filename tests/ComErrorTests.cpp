#include <AppAPI/ComSupport.hpp>
#include <gtest/gtest.h>


static void CheckComErrorType(HRESULT hr1) {
    _com_error err(hr1);

    HRESULT hr2 = err.Error();
    EXPECT_EQ(hr1, hr2);

    // ErrorMessage() is const TCHAR* (wchar_t if _UNICODE, else char), matching comdef.h.
    // Don't check string content: Windows vs. other platforms do not use the same text.
    const TCHAR* msg = err.ErrorMessage();
    EXPECT_TRUE(msg);
}

TEST(ComErrorTests, Test_com_error) {
    // test 3 common error types
    CheckComErrorType(E_INVALIDARG);
    CheckComErrorType(E_BOUNDS);
    CheckComErrorType(E_FAIL);
}
