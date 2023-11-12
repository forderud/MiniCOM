# MiniCOM
Minimal cross-platform implementation of the [Component Object Model (COM)](https://docs.microsoft.com/en-us/windows/win32/com/the-component-object-model) runtime and [Active Template Library (ATL)](https://docs.microsoft.com/en-us/cpp/atl/atl-com-desktop-components).

Designed as a compatibility library to enable building of existing COM/ATL classes with clang or gcc for for Linux, Mac and mobile platforms. Developed due to lack of knowledge of any better alternatives.

### Design goals
* Support simple COM classes implemented in ATL.
* Support most automation-compatible types, so that the same COM classes can be directly accessed from C# and Python (using [comtypes](https://pythonhosted.org/comtypes/)) on Windows without any language wrappers or proxy/stub DLLs for marshaling.
* Compatiblity with clang and gcc.
* Compatibility with Linux, Mac, iOS and Android.

### Missing features
* Complete COM or ATL support.
* Wrapper-code-free access from C# and Python on non-Windows.
* Out-of-process marshalling on non-Windows.

Please contact the author if you're aware of any better alternative, and he'll be happy to scrap this project. I'm hoping that Microsoft [xlang](https://github.com/microsoft/xlang) or a similar project will eventually replace this project, but cross-platform support have so far been postponed.
