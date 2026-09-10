# Runtime shared by every generated binding. Nothing here is interface
# specific, so it is written once instead of being emitted per IDL file.

import ctypes
import struct

# HRESULT values, from the same list as NonWindows.hpp
S_OK          = 0x00000000
E_NOINTERFACE = -0x7FFFBFFE  # 0x80004002
E_POINTER     = -0x7FFFBFFD  # 0x80004003

IID_IUnknown = '00000000-0000-0000-C000-000000000046'


def Succeeded (hr):
    return hr >= 0


def Check (hr, what):
    if not Succeeded(hr):
        raise OSError('%s failed: 0x%08X' % (what, hr & 0xFFFFFFFF))


def Returned (pointer, value):
    '''Write what a handler returned into its [out,retval] argument, if it has one.'''
    if pointer is not None:
        pointer[0] = value
    return S_OK


def GuidBytes (text):
    '''The 16 bytes of a GUID, laid out the way the C struct is.'''
    parts = text.split('-')
    head = struct.pack('<IHH', int(parts[0], 16), int(parts[1], 16), int(parts[2], 16))
    return head + bytes.fromhex(parts[3] + parts[4])


def Call (obj, slot, argtypes, *args):
    '''Call through the object's vtable, which the object pointer points at.'''
    vtable = ctypes.cast(obj, ctypes.POINTER(ctypes.c_void_p)).contents
    entry = ctypes.cast(vtable, ctypes.POINTER(ctypes.c_void_p))[slot]
    signature = ctypes.CFUNCTYPE(ctypes.c_int32, ctypes.c_void_p, *argtypes)
    return signature(entry)(obj, *args)


class Interface:
    '''The IUnknown part of every interface, which is the same for all of them.
       The generated classes add the methods that are not.

       Takes over the reference it is constructed with, the way CComPtr::Attach
       does, and drops it again when garbage collected.'''

    IID = IID_IUnknown

    def __init__ (self, pointer):
        self.pointer = pointer

    def __del__ (self):
        self.Release()

    def QueryInterface (self, interface):
        pointer = ctypes.c_void_p()
        hr = Call(self.pointer, 0, [ctypes.c_char_p, ctypes.POINTER(ctypes.c_void_p)],
                  GuidBytes(interface.IID), ctypes.byref(pointer))
        Check(hr, 'QueryInterface')
        return interface(pointer)

    def AddRef (self):
        return Call(self.pointer, 1, [])

    def Release (self):
        if not self.pointer:
            return 0

        pointer, self.pointer = self.pointer, None
        return Call(pointer, 2, [])


def CreateInstance (library, clsid, interface):
    '''Create a registered COM class, through the library's CoCreateInstance.'''
    library.CoCreateInstance.restype = ctypes.c_int32
    library.CoCreateInstance.argtypes = [ctypes.c_char_p, ctypes.c_void_p, ctypes.c_uint32,
                                         ctypes.c_char_p, ctypes.POINTER(ctypes.c_void_p)]
    pointer = ctypes.c_void_p()
    hr = library.CoCreateInstance(GuidBytes(clsid), None, 0,
                                  GuidBytes(interface.IID), ctypes.byref(pointer))
    Check(hr, 'CoCreateInstance')
    return interface(pointer)


class Implementation:
    '''A COM object backed by a Python object.

    QueryInterface, AddRef and Release are handled here, so that a handler only
    provides the interface methods and cannot get identity or lifetime wrong.
    The generated ImplementXxx functions supply the rest of the method table.
    '''

    def __init__ (self, handler, iids, methods):
        self.handler = handler
        self._iids = [GuidBytes(i) for i in iids]
        self._references = 1

        self._slots = [
            ctypes.CFUNCTYPE(ctypes.c_int32, ctypes.c_void_p, ctypes.c_void_p,
                             ctypes.c_void_p)(self._QueryInterface),
            ctypes.CFUNCTYPE(ctypes.c_uint32, ctypes.c_void_p)(self._AddRef),
            ctypes.CFUNCTYPE(ctypes.c_uint32, ctypes.c_void_p)(self._Release),
        ] + methods

        self._vtable = (ctypes.c_void_p * len(self._slots))(
            *[ctypes.cast(s, ctypes.c_void_p) for s in self._slots])
        self._object = ctypes.pointer(ctypes.cast(self._vtable, ctypes.c_void_p))

    @property
    def pointer (self):
        return ctypes.cast(self._object, ctypes.c_void_p)

    @property
    def references (self):
        return self._references

    def _QueryInterface (self, this, riid, obj):
        if not obj:
            return E_POINTER

        out = ctypes.cast(obj, ctypes.POINTER(ctypes.c_void_p))
        if ctypes.string_at(riid, 16) not in self._iids:
            out[0] = None
            return E_NOINTERFACE

        out[0] = self.pointer.value
        self._references += 1
        return S_OK

    def _AddRef (self, this):
        self._references += 1
        return self._references

    def _Release (self, this):
        self._references -= 1
        return self._references
