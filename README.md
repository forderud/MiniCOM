# MiniCOM
Minimal cross-platform implementation of the [Component Object Model (COM)](https://docs.microsoft.com/en-us/windows/win32/com/the-component-object-model) and [Active Template Library (ATL)](https://docs.microsoft.com/en-us/cpp/atl/atl-com-desktop-components).

Designed as a compatibility library to enable building of existing COM/ATL classes with clang or gcc for for Linux, Mac and mobile platforms. Developed due to lack of knowledge of any better alternatives.

Design goals:
* Support simple COM classes implemented in ATL.
* Support automation-compatible types, so that the same COM classes can be directly accessed from C# and Python (using [comtypes](https://pythonhosted.org/comtypes/)) on Windows without any language wrappers or proxy-stub DLLs.
* Compatiblity with clang and gcc.

NON-goals:
* Complete ATL support
* Out-of-process marshalling on non-Windows (would be nice to have though)
