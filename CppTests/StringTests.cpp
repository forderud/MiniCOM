#include <AppAPI/ComSupport.hpp>
#include <gtest/gtest.h>


TEST(StringTests, Verify_CComBSTR_functionality) {
    // initialization tests
    CComBSTR a(L"a");
    ASSERT_STREQ(a, L"a");

    CComBSTR b1 = L"b";
    ASSERT_STREQ(b1, L"b");

    // copying tests
    CComBSTR b2(b1);
    ASSERT_STREQ(b2, L"b");

    CComBSTR b3 = b1;
    ASSERT_STREQ(b3, L"b");

    // length test
    ASSERT_EQ(a.Length(), 1u);
}

TEST(StringTests, Verify_bstr_t_functionality) {
    // initialization tests
    _bstr_t a(L"a");
    ASSERT_STREQ(a, L"a");

    _bstr_t b1 = L"b";
    ASSERT_STREQ(b1, L"b");

    // copying tests
    _bstr_t b2(b1);
    ASSERT_STREQ(b2, L"b");

    _bstr_t b3 = b1;
    ASSERT_STREQ(b3, L"b");

    // length test
    ASSERT_EQ(a.length(), 1u);
}


TEST(StringTests, Verify_CComBSTR_Concatenation) {
    CComBSTR a(L"a");
    CComBSTR b(L"b");

    // concatenation test
    CComBSTR ab2 = a;
    ab2 += b;
    ASSERT_STREQ(ab2, L"ab");
}

TEST(StringTests, Verify_bstr_t_Concatenation) {
    _bstr_t a(L"a");
    _bstr_t b(L"b");

    // concatenation tests
    _bstr_t ab1 = a + b;
    ASSERT_STREQ(ab1, L"ab");

    _bstr_t ab2 = a;
    ab2 += b;
    ASSERT_STREQ(ab2, L"ab");
}

TEST(StringTests, Verify_CComBSTR_Comparison) {
    CComBSTR str(L"string");

    // compare to a copy of the same string
    ASSERT_TRUE(str == CComBSTR(str));

    // compare different strings
    ASSERT_FALSE(str == CComBSTR(L"another"));
}

TEST(StringTests, Verify_bstr_t_Comparison) {
    _bstr_t str(L"string");

    // compare to a copy of the same string
    ASSERT_TRUE(str == _bstr_t(str));

    // compare different strings
    ASSERT_FALSE(str == _bstr_t(L"another"));
}


TEST(StringTests, Verify_bstr_t_AttachDetach) {
    _bstr_t str1(L"string");

    auto free_str = str1.Detach(); // will leak unless re-attached
    ASSERT_EQ(str1.length(), 0u);

    _bstr_t str2;
    str2.Attach(free_str);
    ASSERT_EQ(str2.length(), 6u);
}

static void PassBSTROutputArgument(/*out*/BSTR* str) {
    assert(str);        // cannot pass nullptr
    ASSERT_FALSE(*str); // string must be empty to prevent leak
    *str = _bstr_t(L"New string").Detach();
}

TEST(StringTests, Verify_BSTR_output_argument) {
    _bstr_t str = L"Initial string";
    // test strings passed as output arguments through COM interfaces using _bstr_t
    PassBSTROutputArgument(str.GetAddress()); // GetAddress will auto-clear existing content
    ASSERT_STREQ(str, L"New string");
}
