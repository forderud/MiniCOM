#include <AppAPI/ComSupport.hpp>
#include <gtest/gtest.h>

#ifndef _WIN32
// these are deliberately not defined in NonWindows.hpp since they easily lead to brittle code
#define VARIANT_TRUE ((VARIANT_BOOL)-1)
#define VARIANT_FALSE ((VARIANT_BOOL)0)
#define TRUE                1
#define FALSE               0
#endif


TEST(BooleanTests, Cast_VARIANT_BOOL_to_bool) {
    VARIANT_BOOL vt = VARIANT_TRUE; // -1
    VARIANT_BOOL vf = VARIANT_FALSE;

    // VARIANT_BOOL is implicitly converted to bool correctly
    EXPECT_TRUE(vt);
    EXPECT_FALSE(vf);

    // need to cast VARIANT_BOOL to bool before comparing to bool
    EXPECT_EQ(static_cast<bool>(vt), true); // would fail if not casting to bool first
    EXPECT_EQ(static_cast<bool>(vf), false);
}

TEST(BooleanTests, Cast_bool_to_VARIANT_BOOL) {
    // implicit casting of bool to VARIANT_BOOL
    VARIANT_BOOL vt = true; // yields 1 instead of VARIANT_TRUE(-1), but that doesn't matter for well-behaved code
    VARIANT_BOOL vf = false; // yields VARIANT_FALSE

    // VARIANT_BOOL is implicitly converted to bool correctly
    EXPECT_TRUE(vt);
    EXPECT_FALSE(vf);

    // need to cast VARIANT_BOOL to bool before comparing to bool
    EXPECT_EQ(static_cast<bool>(vt), true); // would fail if not casting to bool first
    EXPECT_EQ(static_cast<bool>(vf), false);
}


TEST(BooleanTests, Cast_BOOL_to_bool) {
    BOOL bt = TRUE;
    BOOL bf = FALSE;

    // BOOL is implicitly converted to bool correctly
    EXPECT_TRUE(bt);
    EXPECT_FALSE(bf);

    // comparison between BOOL & bool works automatically
    EXPECT_EQ(bt, true);
    EXPECT_EQ(bf, false);
}

TEST(BooleanTests, Cast_bool_to_BOOL) {
    // implicit casting of bool to BOOL
    BOOL bt = true;
    BOOL bf = false;

    // BOOL is implicitly converted to bool correctly
    EXPECT_TRUE(bt);
    EXPECT_FALSE(bf);

    // comparison between BOOL & bool works automatically
    EXPECT_EQ(bt, true);
    EXPECT_EQ(bf, false);
}
