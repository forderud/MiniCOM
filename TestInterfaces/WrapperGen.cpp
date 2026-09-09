// Generate C++ tlh/tli wrapper headers
#ifdef _WIN32
  #ifdef _DEBUG
    #import "TestInterfaces.dir/Debug/Example.tlb"
  #else
    #import "TestInterfaces.dir/Release/Example.tlb"
  #endif
#endif
