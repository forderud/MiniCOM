#include <gtest/gtest.h>
#include <AppAPI/ComSupport.hpp>


#ifdef _WIN32
class AppAPITestsModule :
    public ATL::CAtlExeModuleT<AppAPITestsModule> {
};
AppAPITestsModule _AtlModule;
#endif

int main(int argc, char **argv) {
    testing::InitGoogleTest(&argc, argv);
    return RUN_ALL_TESTS();
}
