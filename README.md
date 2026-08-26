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

## Binding from other languages

`CoCreateInstance` and `CLSIDFromProgID` are provided with the same names and
signatures as on Windows, so the same source creates objects on either platform.
They are `extern "C"`, and the reference parameters are pointers at the ABI
level, so a caller with only a C FFI can use them as they stand:

```python
lib.CLSIDFromProgID("Example", ctypes.byref(clsid))          # -> S_OK
lib.CoCreateInstance(ctypes.byref(clsid), None, CLSCTX_INPROC_SERVER,
                     ctypes.byref(iid), ctypes.byref(obj))   # -> S_OK
```

Two further things are worth knowing before writing such a binding, because both
are invisible from the outside and neither fails in a way that points at itself.

**`IUnknown` is a virtual base, so a secondary interface pointer cannot be used
for reference counting.** `IdlParse.py` emits `struct IFoo : virtual public
IUnknown` — to avoid duplicating `m_ref` under multiple inheritance — and with
virtual inheritance the secondary interface's vtable holds *null* where
`QueryInterface`/`AddRef`/`Release` would be. Those live in the virtual base,
reached through the vtable's virtual-base offset. C++ performs that adjustment
silently; a binding calling through the raw vtable jumps to address zero. The
simple rule is to keep the pointer `CoCreateInstance` returned (or one obtained
by querying for `IID_IUnknown`) and do all lifetime management through it, using
other interface pointers only to call their own methods. Reference counts are
per object rather than per interface, so this is safe.

**Method slots are offset by two relative to Windows.** `IUnknown` here declares
a virtual destructor, which the Itanium ABI gives two vtable slots, so
`QueryInterface` sits at slot 2 and an interface's own methods start at slot 5 —
against 0 and 3 for MSVC COM. A binding targeting both needs that constant.

## Shared & weak references
The repo also contains a [`SharedRef`](SharedRef.hpp) wrapper class for non-owning weak references through a `IWeakRef` interface. This is similar to [`IWeakReference`](https://learn.microsoft.com/en-us/windows/win32/api/weakreference/nn-weakreference-iweakreference), but is also compatible with classical `IUnknown`-based COM.

#### External references 
* Raymond Chen: [Inside STL: Smart pointers](https://devblogs.microsoft.com/oldnewthing/20230814-00/?p=108597) (documents the weak ref-count trick)
* Microsoft: [`_Ref_count_base::_Decref()`](https://github.com/microsoft/STL/blob/main/stl/inc/memory#L1181), [`_Ref_count_base::_Decwref()`](https://github.com/microsoft/STL/blob/main/stl/inc/memory#L1188) implementation (used as inspiration for `SharedRef`)
* LLVM [`__shared_count::__release_shared()`](https://github.com/llvm/llvm-project/blob/main/libcxx/src/memory.cpp#L42), [`__shared_weak_count::__release_weak()`](https://github.com/llvm/llvm-project/blob/main/libcxx/src/memory.cpp#L60) implementation
* GCC [`_Sp_counted_base<_S_atomic>::_M_release()`](https://github.com/gcc-mirror/gcc/blob/master/libstdc%2B%2B-v3/include/bits/shared_ptr_base.h#L392), [`_Sp_counted_base::_M_weak_release()`](https://github.com/gcc-mirror/gcc/blob/master/libstdc%2B%2B-v3/include/bits/shared_ptr_base.h#L213) implementation
