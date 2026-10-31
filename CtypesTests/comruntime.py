# Runtime shared by every generated binding. Nothing here is interface
# specific, so it is written once instead of being emitted per IDL file.
#
# A generated binding declares its types and method signatures, and everything
# that moves a value between Python and C++ is here. Each IDL type is a small
# object with the same four operations, so that an argument, a struct field
# and an array element of that type are handled the same way:
#
#   ctype        its C representation
#   ToC(v)       a C value for a Python one, which owns what it allocates
#   FromC(c, t)  the Python value for a C one, which takes ownership when 't'
#   Free(c)      release what a C value owns
#
# Ownership follows COM: an [in] argument stays the caller's, and whatever is
# handed out through an [out] argument belongs to the receiver.

import ctypes
import enum
import inspect
import struct
import sys
import traceback

# HRESULT values, from the same list as NonWindows.hpp
S_OK          = 0x00000000
E_NOTIMPL     = -0x7FFFBFFF  # 0x80004001
E_NOINTERFACE = -0x7FFFBFFE  # 0x80004002
E_POINTER     = -0x7FFFBFFD  # 0x80004003
E_FAIL        = -0x7FFFBFFB  # 0x80004005

NAMES = {
    0x8000000B: 'E_BOUNDS',
    0x80004001: 'E_NOTIMPL',
    0x80004002: 'E_NOINTERFACE',
    0x80004003: 'E_POINTER',
    0x80004004: 'E_ABORT',
    0x80004005: 'E_FAIL',
    0x8000FFFF: 'E_UNEXPECTED',
    0x80070005: 'E_ACCESSDENIED',
    0x80070006: 'E_HANDLE',
    0x8007000E: 'E_OUTOFMEMORY',
    0x80070057: 'E_INVALIDARG',
    0x80070490: 'E_NOT_SET',
    0x80040154: 'REGDB_E_CLASSNOTREG',
    0x800401F3: 'CO_E_CLASSSTRING',
}

VARIANT_TRUE  = -1
VARIANT_FALSE = 0

IID_IUnknown = '00000000-0000-0000-C000-000000000046'


class COMError (OSError):
    '''A method that returned a failed HRESULT, which is kept in 'hr'.'''

    def __init__ (self, hr, what):
        self.hr = hr & 0xFFFFFFFF
        name = NAMES.get(self.hr, '0x%08X' % self.hr)
        OSError.__init__(self, '%s failed: %s' % (what, name))


def Succeeded (hr):
    return hr >= 0


def Check (hr, what):
    if not Succeeded(hr):
        raise COMError(hr, what)


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


def _Plain (c):
    '''The Python value of a simple C value, which ctypes hands out either as
       a ctypes object or already converted.'''
    if isinstance(c, ctypes._SimpleCData):
        return c.value
    return c


# NonWindows.hpp allocates with malloc, and a BSTR with wcsdup, and frees with free
_libc = ctypes.CDLL(None)
_libc.malloc.restype = ctypes.c_void_p
_libc.malloc.argtypes = [ctypes.c_size_t]
_libc.wcsdup.restype = ctypes.c_void_p
_libc.wcsdup.argtypes = [ctypes.c_wchar_p]
_libc.free.argtypes = [ctypes.c_void_p]


# ---------------------------------------------------------------------- types

class Value:
    '''A base type, such as int or double, which owns nothing.'''

    def __init__ (self, ctype):
        self.ctype = ctype

    def ToC (self, value):
        return self.ctype(value)

    def FromC (self, c, take):
        return _Plain(c)

    def Free (self, c):
        pass

    def Default (self):
        return self.ctype().value


class _Bool (Value):
    '''VARIANT_BOOL, which is VARIANT_TRUE (-1) for true.'''

    def __init__ (self):
        Value.__init__(self, ctypes.c_short)

    def ToC (self, value):
        if value:
            return ctypes.c_short(VARIANT_TRUE)
        return ctypes.c_short(VARIANT_FALSE)

    def FromC (self, c, take):
        return _Plain(c) != VARIANT_FALSE

    def Default (self):
        return False

Bool = _Bool()


class _String:
    '''BSTR, a wcsdup'ed string.'''
    ctype = ctypes.c_void_p

    def ToC (self, text):
        if text is None:
            return ctypes.c_void_p()
        return ctypes.c_void_p(_libc.wcsdup(text))

    def FromC (self, c, take):
        address = _Plain(c)
        if not address:
            return None
        text = ctypes.wstring_at(address)
        if take:
            _libc.free(address)
        return text

    def Free (self, c):
        address = _Plain(c)
        if address:
            _libc.free(address)

    def Default (self):
        return None

String = _String()


class InterfaceOf:
    '''An interface pointer, seen from Python as an instance of its class.'''
    ctype = ctypes.c_void_p

    def __init__ (self, interface):
        self.interface = interface

    def ToC (self, obj):
        '''A pointer with a reference of its own, released by Free.'''
        if obj is None:
            return ctypes.c_void_p()
        address = obj
        if not isinstance(obj, int):
            address = obj._as_parameter_.value
        Call(address, 1, []) # AddRef
        return ctypes.c_void_p(address)

    def FromC (self, c, take):
        address = _Plain(c)
        if not address:
            return None
        if not take:
            Call(address, 1, []) # AddRef, for the reference the wrapper releases
        return self.interface(ctypes.c_void_p(address))

    def Free (self, c):
        address = _Plain(c)
        if address:
            Call(address, 2, []) # Release

    def Default (self):
        return None


class EnumOf:
    '''An enum, which is a 32bit int in C. A value that is no member, such as a
       combination of flags, is seen as a plain int.'''
    ctype = ctypes.c_int

    def __init__ (self, enumeration):
        self.enumeration = enumeration

    def ToC (self, value):
        return ctypes.c_int(int(value))

    def FromC (self, c, take):
        value = _Plain(c)
        try:
            return self.enumeration(value)
        except ValueError:
            return value

    def Free (self, c):
        pass

    def Default (self):
        return self.FromC(0, False)


class FixedArray:
    '''"T name[N]", seen from Python as a list. An argument without a size,
       "T name[]", takes its size from the list passed.'''

    def __init__ (self, element, size):
        self.element = element
        self.size = size

    @property
    def ctype (self):
        return self.element.ctype * self.size

    def ToC (self, values):
        if self.size is not None and len(values) != self.size:
            raise ValueError('expected %d values, got %d' % (self.size, len(values)))
        array = (self.element.ctype * len(values))()
        for i in range(len(values)):
            array[i] = self.element.ToC(values[i])
        return array

    def FromC (self, c, take):
        values = []
        for i in range(self.size):
            values.append(self.element.FromC(c[i], take))
        return values

    def Free (self, c):
        for i in range(len(c)):
            self.element.Free(c[i])

    def Default (self):
        values = []
        for i in range(self.size):
            values.append(self.element.Default())
        return values


class _Buffer (ctypes.Structure):
    '''NonWindows.hpp Buffer<T>: size, pointer and whether it owns the memory.'''
    _fields_ = [('size', ctypes.c_size_t), ('ptr', ctypes.c_void_p), ('owning', ctypes.c_bool)]


class _SafeArray (ctypes.Structure):
    '''NonWindows.hpp SAFEARRAY: one of three Buffers, as its type says.'''
    _fields_ = [
        ('type', ctypes.c_int),
        ('data', _Buffer),
        ('strings', _Buffer),     # Buffer<CComBSTR>
        ('pointers', _Buffer),    # Buffer<CComPtr<IUnknown>>
        ('elm_size', ctypes.c_uint),
    ]

_EMPTY, _DATA, _STRINGS, _POINTERS = 0, 1, 2, 3


class SafeArrayOf:
    '''SAFEARRAY(T), seen from Python as a list, and a SAFEARRAY(BYTE) as bytes.

       The layout is NonWindows.hpp's, which keeps base types in a byte buffer
       and strings and interface pointers each in a buffer of their own.'''
    ctype = ctypes.c_void_p

    def __init__ (self, element):
        self.element = element

    def _Kind (self):
        if self.element is String:
            return _STRINGS
        if isinstance(self.element, InterfaceOf):
            return _POINTERS
        return _DATA

    def _IsBytes (self):
        return isinstance(self.element, Value) and self.element.ctype == ctypes.c_ubyte

    def ToC (self, values):
        if values is None:
            return ctypes.c_void_p()
        kind = self._Kind()
        address = _libc.malloc(ctypes.sizeof(_SafeArray))
        ctypes.memset(address, 0, ctypes.sizeof(_SafeArray))
        array = _SafeArray.from_address(address)
        array.type = kind
        array.data.owning = array.strings.owning = array.pointers.owning = True

        if kind == _DATA:
            array.elm_size = ctypes.sizeof(self.element.ctype)
            array.data.size = len(values)*array.elm_size
            if array.data.size:
                array.data.ptr = _libc.malloc(array.data.size)
                if self._IsBytes():
                    ctypes.memmove(array.data.ptr, bytes(values), array.data.size)
                else:
                    elements = (self.element.ctype * len(values)).from_address(array.data.ptr)
                    for i in range(len(values)):
                        elements[i] = self.element.ToC(values[i])
        else:
            array.elm_size = ctypes.sizeof(ctypes.c_void_p)
            buffer = array.strings
            if kind == _POINTERS:
                buffer = array.pointers
            buffer.size = len(values)
            if values:
                buffer.ptr = _libc.malloc(len(values)*ctypes.sizeof(ctypes.c_void_p))
                elements = (ctypes.c_void_p * len(values)).from_address(buffer.ptr)
                for i in range(len(values)):
                    elements[i] = self.element.ToC(values[i]).value
        return ctypes.c_void_p(address)

    def FromC (self, c, take):
        address = _Plain(c)
        if not address:
            return None
        array = _SafeArray.from_address(address)

        values = []
        if array.type == _DATA and array.data.ptr:
            if self._IsBytes():
                values = ctypes.string_at(array.data.ptr, array.data.size)
            else:
                count = array.data.size // array.elm_size
                elements = (self.element.ctype * count).from_address(array.data.ptr)
                for i in range(count):
                    values.append(self.element.FromC(elements[i], False))
        elif array.type in [_STRINGS, _POINTERS]:
            buffer = array.strings
            if array.type == _POINTERS:
                buffer = array.pointers
            if buffer.ptr:
                elements = (ctypes.c_void_p * buffer.size).from_address(buffer.ptr)
                for i in range(buffer.size):
                    values.append(self.element.FromC(elements[i], False))
        elif self._IsBytes():
            values = b''

        if take:
            self.Free(c)
        return values

    def Free (self, c):
        '''What SAFEARRAY::Destroy does.'''
        address = _Plain(c)
        if not address:
            return
        array = _SafeArray.from_address(address)
        for buffer, element in [(array.strings, String), (array.pointers, InterfaceOf(None))]:
            if buffer.ptr and buffer.owning:
                elements = (ctypes.c_void_p * buffer.size).from_address(buffer.ptr)
                for i in range(buffer.size):
                    element.Free(elements[i])
        for buffer in [array.data, array.strings, array.pointers]:
            if buffer.ptr and buffer.owning:
                _libc.free(buffer.ptr)
        _libc.free(address)

    def Default (self):
        if self._IsBytes():
            return b''
        return []


class Struct:
    '''Base of the generated struct classes, whose FIELDS list (name, type).
       An instance holds the fields as Python values, and is constructed from
       keyword arguments, with the rest of the fields at their default.'''

    FIELDS = []
    _layout = None

    def __init__ (self, **values):
        for name, type in self.FIELDS:
            if name in values:
                setattr(self, name, values.pop(name))
            else:
                setattr(self, name, type.Default())
        if values:
            raise TypeError('%s has no field %s' % (self.__class__.__name__, ', '.join(values)))

    def __eq__ (self, other):
        if type(other) is not type(self):
            return NotImplemented
        for name, _ in self.FIELDS:
            if getattr(self, name) != getattr(other, name):
                return False
        return True

    def __repr__ (self):
        fields = []
        for name, _ in self.FIELDS:
            fields.append('%s=%r' % (name, getattr(self, name)))
        return '%s(%s)' % (self.__class__.__name__, ', '.join(fields))

    @classmethod
    def Layout (cls):
        '''The C struct, laid out by ctypes the way the C compiler does.'''
        if cls.__dict__.get('UNSUPPORTED'):
            raise NotImplementedError('%s: %s' % (cls.__name__, cls.UNSUPPORTED))
        if cls.__dict__.get('_layout') is None:
            fields = []
            for name, field in cls.FIELDS:
                fields.append((name, field.ctype))
            cls._layout = type(cls.__name__+'Layout', (ctypes.Structure,), {'_fields_': fields})
        return cls._layout


class StructOf:
    '''A struct, seen from Python as an instance of its generated class.'''

    def __init__ (self, record):
        self.record = record

    @property
    def ctype (self):
        return self.record.Layout()

    def ToC (self, value):
        c = self.ctype()
        for name, type in self.record.FIELDS:
            setattr(c, name, type.ToC(getattr(value, name)))
        return c

    def FromC (self, c, take):
        value = self.record.__new__(self.record)
        for name, type in self.record.FIELDS:
            setattr(value, name, type.FromC(getattr(c, name), take))
        return value

    def Free (self, c):
        for name, type in self.record.FIELDS:
            type.Free(getattr(c, name))

    def Default (self):
        return self.record()


class Enum (enum.IntEnum):
    '''Base of the generated enum classes.'''


def Of (cls):
    '''The type of a generated struct or enum class.'''
    if issubclass(cls, Struct):
        return StructOf(cls)
    return EnumOf(cls)


# ----------------------------------------------------------------- arguments

IN, OUT, INOUT = 'in', 'out', 'in,out'


class Param:
    '''One argument of a method: its direction, its type, and how the vtable
       passes it, which is by value, as a pointer to the value, or as a pointer
       to the first element of an array.'''

    def __init__ (self, name, direction, type, passing):
        self.name = name
        self.direction = direction
        self.type = type
        self.passing = passing

    def ArgType (self):
        if self.passing == 'value':
            return self.type.ctype
        if self.passing == 'array':
            return ctypes.POINTER(self.type.element.ctype)
        return ctypes.POINTER(self.type.ctype)


def In (name, type):
    return Param(name, IN, type, 'value')

def InPointer (name, type):
    return Param(name, IN, type, 'pointer')

def InArray (name, type):
    return Param(name, IN, type, 'array')

def Out (name, type):
    return Param(name, OUT, type, 'pointer')

def OutArray (name, type):
    return Param(name, OUT, type, 'array')

def InOut (name, type):
    return Param(name, INOUT, type, 'pointer')


def _Store (pointer, index, c):
    '''Write a C value to where a pointer points.'''
    if isinstance(c, ctypes._SimpleCData):
        pointer[index] = c.value
    else:
        pointer[index] = c


class Method:
    '''A method in an interface's vtable. A caller passes the [in] arguments and
       gets the [out] ones back: nothing when there are none, the value when
       there is one, and a tuple in declaration order when there are several,
       the way comtypes does it. A Python handler implementing the method gets
       and returns the same.'''

    def __init__ (self, name, *params):
        self.name = name
        self.params = params
        self.slot = None
        self.owner = None

    def Inputs (self):
        names = []
        for param in self.params:
            if param.direction != OUT:
                names.append(param.name)
        return names

    def ArgTypes (self):
        types = []
        for param in self.params:
            types.append(param.ArgType())
        return types

    def Invoke (self, obj, args):
        what = self.owner.__name__+'.'+self.name
        if len(args) != len(self.Inputs()):
            raise TypeError('%s takes %d arguments (%s), got %d' % (what, len(self.Inputs()), ', '.join(self.Inputs()), len(args)))

        values = []  # what the vtable is passed
        inputs = []  # [in] arguments, which stay the caller's
        outputs = [] # where the [out] arguments are written
        i = 0
        try:
            for param in self.params:
                if param.direction == OUT:
                    c = param.type.ctype()
                else:
                    c = param.type.ToC(args[i])
                    i += 1

                if param.direction == IN:
                    inputs.append((param, c))
                else:
                    outputs.append((param, c))

                if param.passing == 'pointer':
                    values.append(ctypes.byref(c))
                else:
                    values.append(c)

            hr = Call(obj.pointer, self.slot, self.ArgTypes(), *values)
        except BaseException:
            for param, c in outputs:
                if param.direction == INOUT:
                    param.type.Free(c) # converted, but never passed
            raise
        finally:
            for param, c in inputs:
                param.type.Free(c)

        if not Succeeded(hr):
            for param, c in outputs:
                param.type.Free(c)
            Check(hr, what)

        results = []
        for param, c in outputs:
            results.append(param.type.FromC(c, True))
        if not results:
            return None
        if len(results) == 1:
            return results[0]
        return tuple(results)

    def Thunk (self, handler):
        '''The vtable entry for a Python handler. Nothing may escape back into
           C++: ctypes would turn an exception into an arbitrary return value
           that can look like success.'''
        def Entry (this, *raw):
            method = getattr(handler, self.name, None)
            if method is None:
                return E_NOTIMPL

            try:
                args = []
                outputs = []
                for i in range(len(self.params)):
                    param, value = self.params[i], raw[i]
                    if param.direction != IN and not value:
                        return E_POINTER
                    if param.passing == 'array' and param.direction == IN:
                        if param.type.size is None:
                            args.append(value) # the size is not known, so pass the pointer
                        else:
                            args.append(param.type.FromC(value, False))
                        continue
                    if param.passing == 'pointer':
                        if not value: # a null [in] pointer
                            args.append(None)
                            continue
                        value = value[0]
                    if param.direction != OUT:
                        args.append(param.type.FromC(value, False))
                    if param.direction != IN:
                        outputs.append((param, raw[i]))

                result = method(*args)

                results = [result]
                if len(outputs) > 1:
                    if not isinstance(result, tuple) or len(result) != len(outputs):
                        raise TypeError('%s must return a tuple of %d values' % (self.name, len(outputs)))
                    results = result
                for j in range(len(outputs)):
                    param, pointer = outputs[j]
                    c = param.type.ToC(results[j])
                    if param.passing == 'array':
                        for k in range(param.type.size):
                            _Store(pointer, k, c[k])
                    else:
                        if param.direction == INOUT:
                            param.type.Free(pointer[0]) # the value passed in is replaced
                        _Store(pointer, 0, c)
            except Exception:
                sys.stderr.write('%s.%s raised\n' % (type(handler).__name__, self.name))
                traceback.print_exc()
                return E_FAIL
            return S_OK

        return ctypes.CFUNCTYPE(ctypes.c_int32, ctypes.c_void_p, *self.ArgTypes())(Entry)

    def Bind (self):
        '''The method as the interface class has it.'''
        method = self
        def Bound (obj, *args):
            return method.Invoke(obj, args)
        Bound.__name__ = self.name
        parameters = [inspect.Parameter('self', inspect.Parameter.POSITIONAL_ONLY)]
        for name in self.Inputs():
            parameters.append(inspect.Parameter(name, inspect.Parameter.POSITIONAL_ONLY))
        Bound.__signature__ = inspect.Signature(parameters)
        return Bound


class Unsupported:
    '''A method with an argument that the bindings cannot pass. It keeps its
       slot, so that the other methods are unaffected, but raises when called,
       and answers E_NOTIMPL when implemented in Python.'''

    def __init__ (self, name, reason):
        self.name = name
        self.reason = reason
        self.slot = None
        self.owner = None

    def Thunk (self, handler):
        return ctypes.CFUNCTYPE(ctypes.c_int32, ctypes.c_void_p)(lambda this, *args: E_NOTIMPL)

    def Bind (self):
        method = self
        def Bound (obj, *args):
            raise NotImplementedError('%s.%s: %s' % (method.owner.__name__, method.name, method.reason))
        Bound.__name__ = self.name
        return Bound


# ----------------------------------------------------------------- interfaces

class Interface:
    '''The IUnknown part of every interface, which is the same for all of them.
       The generated classes add the methods that are not.

       Takes over the reference it is constructed with, the way CComPtr::Attach
       does, and drops it again when garbage collected.'''

    IID = IID_IUnknown
    VTABLE = [] # the methods after IUnknown's three, inherited ones first

    def __init__ (self, pointer):
        self.pointer = pointer

    def __del__ (self):
        self.Release()

    @property
    def _as_parameter_ (self):
        '''Lets the object itself be passed where ctypes expects the pointer.'''
        return self.pointer

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


def Define (interface, methods):
    '''Give a generated interface class its methods, which follow its base's in
       the vtable. The base is defined first, even when it is declared in
       another IDL file, so the slots follow from the class hierarchy.'''
    base = interface.__mro__[1]
    first = 3+len(base.VTABLE) # after QueryInterface, AddRef and Release
    interface.VTABLE = base.VTABLE+methods
    for i in range(len(methods)):
        methods[i].slot = first+i
        methods[i].owner = interface
        setattr(interface, methods[i].name, methods[i].Bind())


def CreateInstance (library, clsid, interface):
    '''Create a registered COM class, through the library's CoCreateInstance.'''
    return _Create(library, GuidBytes(clsid), interface)


def CreateByName (library, progid, interface):
    '''Create a registered COM class by its ProgID, "<Program>.<Class>.<Version>"
       or just the class name, through the library's CLSIDFromProgID. The IDL of
       an interface library need not declare its classes, so this is how
       classes without a coclass in the IDL are reached.'''
    library.CLSIDFromProgID.restype = ctypes.c_int32
    library.CLSIDFromProgID.argtypes = [ctypes.c_wchar_p, ctypes.c_void_p]
    clsid = ctypes.create_string_buffer(16)
    Check(library.CLSIDFromProgID(progid, clsid), 'CLSIDFromProgID')
    return _Create(library, clsid.raw, interface)


def _Create (library, clsid, interface):
    library.CoCreateInstance.restype = ctypes.c_int32
    library.CoCreateInstance.argtypes = [ctypes.c_char_p, ctypes.c_void_p, ctypes.c_uint32,
                                         ctypes.c_char_p, ctypes.POINTER(ctypes.c_void_p)]
    pointer = ctypes.c_void_p()
    hr = library.CoCreateInstance(clsid, None, 0, GuidBytes(interface.IID), ctypes.byref(pointer))
    Check(hr, 'CoCreateInstance')
    return interface(pointer)


# Implementations that the library holds a reference to. The library only has
# their address, so without this they could be collected while it still calls them.
_held = set()


class Implementation:
    '''A COM object backed by a Python object, which implements an interface
       class and everything it inherits from.

    QueryInterface, AddRef and Release are handled here, so that a handler only
    provides the interface methods and cannot get identity or lifetime wrong.

    The Python object holds the first reference, and stays alive for as long as
    the library holds any of the others, even when Python has let go of it.
    '''

    def __init__ (self, handler, interface):
        self.handler = handler
        self._iids = []
        for cls in interface.__mro__:
            if issubclass(cls, Interface):
                self._iids.append(GuidBytes(cls.IID))
        self._references = 1

        self._slots = [
            ctypes.CFUNCTYPE(ctypes.c_int32, ctypes.c_void_p, ctypes.c_void_p,
                             ctypes.c_void_p)(self._QueryInterface),
            ctypes.CFUNCTYPE(ctypes.c_uint32, ctypes.c_void_p)(self._AddRef),
            ctypes.CFUNCTYPE(ctypes.c_uint32, ctypes.c_void_p)(self._Release),
        ]
        for method in interface.VTABLE:
            self._slots.append(method.Thunk(handler))

        self._vtable = (ctypes.c_void_p * len(self._slots))()
        for i in range(len(self._slots)):
            self._vtable[i] = ctypes.cast(self._slots[i], ctypes.c_void_p)
        self._object = ctypes.pointer(ctypes.cast(self._vtable, ctypes.c_void_p))

    @property
    def pointer (self):
        return ctypes.cast(self._object, ctypes.c_void_p)

    @property
    def _as_parameter_ (self):
        '''Lets the object itself be passed where ctypes expects the pointer.'''
        return self.pointer

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
        self._Count(+1)
        return S_OK

    def _AddRef (self, this):
        return self._Count(+1)

    def _Release (self, this):
        return self._Count(-1)

    def _Count (self, change):
        self._references += change
        if self._references > 1:
            _held.add(self)
        else:
            _held.discard(self)
        return self._references
