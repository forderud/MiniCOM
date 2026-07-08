Partial cross-platform implementation of the [Component Object Model (COM)](https://docs.microsoft.com/en-us/windows/win32/com/the-component-object-model) runtime and [Active Template Library (ATL)](https://docs.microsoft.com/en-us/cpp/atl/atl-com-desktop-components).

Designed as a **compatibility library to enable usage of existing COM/ATL classes also on non-Windows platforms**.

Developed due to lack of knowledge of any better alternatives. Please contact the author if you're aware of any better alternative, and he'll be happy to scrap this project. I'm hoping that Microsoft [xlang](https://github.com/microsoft/xlang) or a similar project will eventually replace this project, but cross-platform support have so far been postponed.

### Design goals
* Support most COM classes implemented in ATL.
* Support most automation-compatible types, so that the same COM classes can be directly accessed from C# and Python (using [comtypes](https://pythonhosted.org/comtypes/)) on Windows without any language wrappers or proxy/stub DLLs for marshaling.

### Platform support
The following operating systems are currently supported:
* Linux
* MacOS
* Android
* iOS
* WebAssembly with [Emscripten](https://emscripten.org/) compiler

Both the [gcc](https://gcc.gnu.org/) and [clang](https://clang.llvm.org/) compilers are supported.

There's no point in supporting Windows, since the same functionality is already inbuilt there.

### Missing features
* Complete COM or ATL support.
* Wrapper-code-free access from C# and Python on non-Windows.
* Out-of-process marshalling on non-Windows.

Contributions for addressing missing features are welcome.

## Shared & weak references
The repo also contains a [`SharedRef`](SharedRef.hpp) wrapper class for non-owning weak references through a `IWeakRef` interface. This is similar to [`IWeakReference`](https://learn.microsoft.com/en-us/windows/win32/api/weakreference/nn-weakreference-iweakreference), but is also compatible with classical `IUnknown`-based COM.

### External references 
* Raymond Chen: [Inside STL: Smart pointers](https://devblogs.microsoft.com/oldnewthing/20230814-00/?p=108597)
* LLVM [`__shared_weak_count::__release_weak()`](https://github.com/llvm/llvm-project/blob/main/libcxx/src/memory.cpp#L60) implementation
* GCC [`_Sp_counted_base<_S_single>::_M_release()`](https://github.com/gcc-mirror/gcc/blob/master/libstdc%2B%2B-v3/include/bits/shared_ptr_base.h#L369), [`_Sp_counted_base::_M_weak_release()`](https://github.com/gcc-mirror/gcc/blob/master/libstdc%2B%2B-v3/include/bits/shared_ptr_base.h#L213) implementation

