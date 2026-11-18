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


TEST(StringTests, Verify_BSTR_memory_layout) {
    // allocate a BSTR with an embedded zero
    const wchar_t src[] = {L'a', L'b', L'\0', L'c'};
    BSTR str = SysAllocStringLen(src, 4);
    ASSERT_NE(str, nullptr);

    // length is stored in bytes in the 4 bytes before the string (sizeof(wchar_t) per character)
    ASSERT_EQ(SysStringLen(str), 4u);
    ASSERT_EQ(SysStringByteLen(str), 4*sizeof(wchar_t));
    ASSERT_EQ(reinterpret_cast<const uint32_t*>(str)[-1], 4*sizeof(wchar_t));

    // characters are copied verbatim and zero terminated
    for (size_t i = 0; i < 4; ++i)
        ASSERT_EQ(str[i], src[i]);
    ASSERT_EQ(str[4], L'\0');

    SysFreeString(str);

    // nullptr is a valid empty string
    ASSERT_EQ(SysAllocString(nullptr), nullptr);
    ASSERT_EQ(SysStringLen(nullptr), 0u);
    ASSERT_EQ(SysStringByteLen(nullptr), 0u);
    SysFreeString(nullptr);
}

TEST(StringTests, Verify_CComBSTR_size_argument) {
    // only the first "size" characters are copied
    CComBSTR str(2, L"abcdef");
    ASSERT_STREQ(str, L"ab");
    ASSERT_EQ(str.Length(), 2u);
    ASSERT_EQ(str.ByteLength(), 2*sizeof(wchar_t));

    // zero size yields an empty (nullptr) string
    CComBSTR empty(0, L"abc");
    ASSERT_EQ(empty.m_str, nullptr);
    ASSERT_EQ(empty.Length(), 0u);
}

TEST(StringTests, Verify_CComBSTR_AttachDetach) {
    CComBSTR str1(L"string");

    BSTR raw = str1.Detach(); // ownership transferred to caller
    ASSERT_EQ(str1.m_str, nullptr);
    ASSERT_EQ(SysStringLen(raw), 6u);

    CComBSTR str2;
    str2.Attach(raw); // ownership transferred back
    ASSERT_STREQ(str2, L"string");

    // Copy returns a BSTR that must be freed with SysFreeString
    BSTR copy = str2.Copy();
    ASSERT_NE(copy, raw);
    ASSERT_STREQ(copy, L"string");
    SysFreeString(copy);
}

TEST(StringTests, Verify_bstr_t_copy_keeps_embedded_zeros) {
    const wchar_t src[] = {L'a', L'\0', L'b'};
    _bstr_t str1(SysAllocStringLen(src, 3), false); // take over ownership

    _bstr_t str2(str1); // copy ctor
    ASSERT_EQ(str2.length(), 3u);

    _bstr_t str3;
    str3 = str1; // copy assignment
    ASSERT_EQ(str3.length(), 3u);

    _bstr_t str4(static_cast<BSTR>(str1), true); // copy from BSTR
    ASSERT_EQ(str4.length(), 3u);
    ASSERT_EQ(static_cast<wchar_t*>(str4)[2], L'b');
}

#ifndef _WIN32
TEST(StringTests, Verify_SysAllocStringLen_overflow) {
    // byte length doesn't fit in the 32bit length prefix
    ASSERT_EQ(SysAllocStringLen(nullptr, 0x40000000u), nullptr);
}
#endif
