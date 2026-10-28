# Generates ctypes bindings from a .winmd file, the metadata format of
# https://github.com/microsoft/win32metadata, which Winmd/TestInterfaces.proj
# generates from the IDL on Windows.
#
# Nothing here parses IDL. The .winmd is read with dnfile, a reader for the
# ECMA-335 tables that .NET assemblies and .winmd files share, and each type is
# turned into a small dict in the shape of win32json, the JSON translation of
# win32metadata. The bindings only declare the types: an enum class, a struct
# class with its fields, and an interface class with the parameters of each
# method. Moving values between Python and C++ is left to comruntime.py, which
# is the same for every IDL file.
#
# A .winmd describes the Windows ABI, and some of what the C++ side is
# compiled with off Windows is not in it. Where that matters, it is said below.

import hashlib
import keyword
import os
import struct
import sys
import uuid

import dnfile

# mirrors the typedefs in NonWindows.hpp. A ctypes type has the size that the
# C compiler gives the C type it is named after on the platform it runs on,
# which is also what the C++ side of the bindings is compiled with.
SCALAR = {
    'char':               'ctypes.c_char',
    'wchar_t':            'ctypes.c_wchar',
    'short':              'ctypes.c_short',
    'int':                'ctypes.c_int',
    'long long':          'ctypes.c_longlong',
    'unsigned char':      'ctypes.c_ubyte',
    'unsigned short':     'ctypes.c_ushort',
    'unsigned int':       'ctypes.c_uint',
    'unsigned long long': 'ctypes.c_ulonglong',
    'float':              'ctypes.c_float',
    'double':             'ctypes.c_double',
    # a typedef from Windows.Win32.winmd keeps its name, so BOOL, which
    # NonWindows.hpp declares as "long", can still be given the size of a C
    # long. A plain "long" cannot: the .winmd only says Int32, which is what
    # long is on Windows, but not on 64bit Linux and macOS.
    'BOOL':               'ctypes.c_long',
    'HRESULT':            'ctypes.c_int32',
    'HWND':               'ctypes.c_void_p',
}

# ECMA-335 II.23.1.16 element types, as the C types the .winmd records them for
ELEMENT = {
    0x01: 'void',
    0x03: 'wchar_t',
    0x04: 'char',
    0x05: 'unsigned char',
    0x06: 'short',
    0x07: 'unsigned short',
    0x08: 'int',
    0x09: 'unsigned int',
    0x0a: 'long long',
    0x0b: 'unsigned long long',
    0x0c: 'float',
    0x0d: 'double',
}
PTR, VALUETYPE, CLASS, ARRAY, SZARRAY, CMOD_REQD, CMOD_OPT = 0x0f, 0x11, 0x12, 0x14, 0x1d, 0x1f, 0x20

# Windows.Win32 typedefs that the bindings pass by name
NATIVE = ['BSTR', 'VARIANT_BOOL', 'BOOL', 'HRESULT', 'HWND']


class Unsupported (Exception):
    '''A type that the bindings cannot pass yet.'''


def PythonName (name):
    '''A name that is safe as a Python identifier in the generated code.'''
    if keyword.iskeyword(name) or name in ['self', 'ctypes', 'comruntime']:
        return name+'_'
    return name


def IsInterface (type):
    return type['Kind'] == 'ApiRef' and type['TargetKind'] == 'Com'


class Blob:
    '''Reads a signature or custom attribute blob front to back.'''

    def __init__ (self, data):
        self.data = bytes(data)
        self.pos = 0

    def Byte (self):
        self.pos += 1
        return self.data[self.pos-1]

    def Peek (self):
        return self.data[self.pos]

    def Compressed (self):
        '''An unsigned integer compressed as in ECMA-335 II.23.2.'''
        first = self.Byte()
        if first & 0x80 == 0:
            return first
        if first & 0xc0 == 0x80:
            return (first & 0x3f) << 8 | self.Byte()
        value = first & 0x1f
        for _ in range(3):
            value = value << 8 | self.Byte()
        return value

    def Unpack (self, format):
        values = struct.unpack_from('<'+format, self.data, self.pos)
        self.pos += struct.calcsize('<'+format)
        return values

    def String (self):
        '''A SerString, as custom attribute arguments are stored.'''
        length = self.Compressed()
        self.pos += length
        return self.data[self.pos-length:self.pos].decode('utf-8')


class Winmd:
    '''The types in one .winmd file, as win32json-shaped dicts.'''

    def __init__ (self, path):
        self.md = dnfile.dnPE(path).net.mdtables

        # custom attributes by what they are on, as (table, row) -> [(name, blob)]
        self.attributes = {}
        for attribute in self.Table('CustomAttribute'):
            constructor = attribute.Type.row
            if attribute.Type.table.name == 'MemberRef':
                name = str(constructor.Class.row.TypeName)
            else:
                name = '?'
            key = (attribute.Parent.table.name, attribute.Parent.row_index)
            self.attributes.setdefault(key, []).append((name, attribute.Value.value))

        self.constants = {} # field row -> value of an enum member
        for constant in self.Table('Constant'):
            self.constants[constant.Parent.row_index] = self.Constant(constant.Type, constant.Value.value)

        self.bases = {} # typedef row -> base interface
        for implementation in self.Table('InterfaceImpl'):
            self.bases[implementation.Class.row_index] = implementation.Interface

        self.typedefs = {} # (namespace, name) -> row, for the TypeRefs to them
        for row in self.Table('TypeDef'):
            self.typedefs[(str(row.TypeNamespace), str(row.TypeName))] = row

        self.nested = set() # typedef rows of nested types, such as fixed buffers
        for nested in self.Table('NestedClass'):
            self.nested.add(nested.NestedClass.row_index)

    def Table (self, name):
        '''The rows of a table, of which a .winmd may have none.'''
        return getattr(self.md, name, None) or []

    @staticmethod
    def Constant (element, data):
        format = {0x04: 'b', 0x05: 'B', 0x06: 'h', 0x07: 'H', 0x08: 'i', 0x09: 'I', 0x0a: 'q', 0x0b: 'Q'}[element]
        return struct.unpack('<'+format, bytes(data))[0]

    def Attributes (self, table, row):
        return self.attributes.get((table, row), [])

    def Guid (self, table, row):
        for name, data in self.Attributes(table, row):
            if name == 'GuidAttribute':
                fields = Blob(data)
                fields.Unpack('H') # prolog
                return str(uuid.UUID(bytes_le=fields.data[2:18])).upper()
        return None

    def Reference (self, index):
        '''A type that a signature or an InterfaceImpl row refers to.'''
        row = index.row
        name = str(row.TypeName)
        namespace = str(row.TypeNamespace)
        if index.table.name == 'TypeRef' and (namespace, name) in self.typedefs:
            row = self.typedefs[(namespace, name)] # declared in this file, and referred to by name
            index = _Index(self.md.TypeDef, 0, row)
        if index.table.name == 'TypeDef':
            com = row.Flags.tdInterface
            return {'Kind': 'ApiRef', 'Name': name, 'Api': namespace, 'TargetKind': 'Com' if com else 'Default'}
        if namespace.startswith('Windows.Win32.'):
            if name == 'IUnknown':
                return {'Kind': 'ApiRef', 'Name': name, 'TargetKind': 'Com'} # no Api: comruntime.Interface
            if name == 'SAFEARRAY':
                # MIDL writes SAFEARRAY(float) into the header as "SAFEARRAY *", so
                # the element type never reaches the .winmd
                return {'Kind': 'SafeArray', 'Child': None}
            if name in NATIVE:
                return {'Kind': 'Native', 'Name': name}
        return {'Kind': 'Native', 'Name': namespace+'.'+name}

    def Token (self, blob):
        '''A TypeDefOrRef coded index in a signature (II.23.2.8).'''
        value = blob.Compressed()
        table = [self.md.TypeDef, self.md.TypeRef, self.md.TypeSpec][value & 3]
        row = table[(value >> 2) - 1]
        return self.Reference(_Index(table, value >> 2, row))

    def Type (self, blob):
        '''The type at the start of a signature blob, as a win32json type tree.'''
        element = blob.Byte()
        while element in (CMOD_REQD, CMOD_OPT): # e.g. "const"
            blob.Compressed()
            element = blob.Byte()
        if element in ELEMENT:
            return {'Kind': 'Native', 'Name': ELEMENT[element]}
        if element == PTR:
            return {'Kind': 'PointerTo', 'Child': self.Type(blob)}
        if element in (VALUETYPE, CLASS):
            type = self.Token(blob)
            if IsInterface(type):
                # an interface is a reference type, so "ICalcCb *" in C is
                # written as just ICalcCb, and "ICalc2 **" as ICalc2 *
                return {'Kind': 'PointerTo', 'Child': type}
            return type
        if element == ARRAY: # II.23.2.13: type, rank, sizes, lower bounds
            child = self.Type(blob)
            rank = blob.Compressed()
            sizes = [blob.Compressed() for _ in range(blob.Compressed())]
            for _ in range(blob.Compressed()):
                blob.Compressed()
            if rank != 1 or len(sizes) != 1:
                raise Unsupported('array of rank '+str(rank))
            return {'Kind': 'Array', 'Shape': {'Size': sizes[0]}, 'Child': child}
        if element == SZARRAY:
            return {'Kind': 'Array', 'Shape': None, 'Child': self.Type(blob)}
        return {'Kind': 'Native', 'Name': 'element type 0x%02x' % element}

    def NativeArrayCount (self, row):
        '''The CountConst of a NativeArrayInfo attribute on a parameter, which is
           how a fixed-size array argument such as "float values[3]" is kept.'''
        for name, data in self.Attributes('Param', row):
            if name != 'NativeArrayInfoAttribute':
                continue
            blob = Blob(data)
            blob.Unpack('H') # prolog, and no fixed arguments
            for _ in range(blob.Unpack('H')[0]):
                blob.Byte() # field or property
                kind = blob.Byte()
                key = blob.String()
                value = blob.Unpack({0x08: 'i', 0x06: 'h'}[kind])[0]
                if key == 'CountConst':
                    return value
        return None

    def Method (self, row):
        blob = Blob(row.Signature.value)
        blob.Byte() # calling convention, HASTHIS
        count = blob.Compressed()
        self.Type(blob) # the HRESULT
        types = [self.Type(blob) for _ in range(count)]

        params = []
        for index in row.ParamList:
            param = index.row
            if param.Sequence == 0:
                continue # the return value
            attrs = []
            if param.Flags.pdIn:
                attrs.append('In')
            if param.Flags.pdOut:
                attrs.append('Out')
            names = [name for name, _ in self.Attributes('Param', index.row_index)]
            if 'RetValAttribute' in names:
                attrs.append('RetVal')
            type = types[param.Sequence-1]
            count = self.NativeArrayCount(index.row_index)
            if count is not None and type['Kind'] == 'PointerTo':
                type = {'Kind': 'Array', 'Shape': {'Size': count}, 'Child': type['Child']}
            params.append({'Name': str(param.Name), 'Type': type, 'Attrs': attrs})
        return {'Name': str(row.Name), 'Params': params}

    def Types (self, namespace):
        types = []
        for number, row in enumerate(self.md.TypeDef, 1):
            if str(row.TypeNamespace) != namespace or number in self.nested:
                continue
            name = str(row.TypeName)
            guid = self.Guid('TypeDef', number)

            if row.Flags.tdInterface:
                base = {'Kind': 'ApiRef', 'Name': 'IUnknown', 'TargetKind': 'Com'}
                if number in self.bases:
                    base = self.Reference(self.bases[number])
                methods = [self.Method(index.row) for index in row.MethodList]
                types.append({'Kind': 'Com', 'Name': name, 'Guid': guid, 'Interface': base, 'Methods': methods})
                continue

            extends = str(row.Extends.row.TypeName) if row.Extends.row_index else ''
            fields = [(index.row_index, index.row) for index in row.FieldList]
            if extends == 'Enum':
                values = []
                for field_number, field in fields:
                    if field_number in self.constants:
                        values.append({'Name': str(field.Name), 'Value': self.constants[field_number]})
                types.append({'Kind': 'Enum', 'Name': name, 'Values': values})
            elif extends == 'ValueType' and not fields and guid:
                # a coclass, of which only the CLSID is kept, and not the
                # interfaces it implements
                types.append({'Kind': 'ComClassID', 'Name': name, 'Guid': guid})
            elif extends == 'ValueType':
                struct_fields = []
                for _, field in fields:
                    blob = Blob(field.Signature.value)
                    blob.Byte() # FIELD
                    struct_fields.append({'Name': str(field.Name), 'Type': self.Type(blob)})
                types.append({'Kind': 'Struct', 'Name': name, 'Fields': struct_fields})
            # else the Apis class, which holds the IID and CLSID constants
        return types


class _Index:
    '''A coded index made from a signature token, shaped like dnfile's.'''
    def __init__ (self, table, row_index, row):
        self.table = table
        self.row_index = row_index
        self.row = row


class Bindings:
    '''Generates the bindings for the types of one namespace, which is one IDL
       file. A type from another namespace is taken from the bindings for it,
       which are imported as <namespace>Bindings.'''

    def __init__ (self, winmd, namespace):
        self.api = namespace
        self.types = Winmd(winmd).Types(namespace)
        if not self.types:
            raise Exception('no types in namespace "'+namespace+'" of '+winmd)
        self.imports = []     # other files whose bindings are referred to
        self.unsupported = [] # what is left out, with the reason

    def ClassName (self, reference):
        '''The Python class for a type that the .winmd refers to by name.'''
        if 'Api' not in reference:
            if IsInterface(reference):
                return 'comruntime.Interface' # IUnknown
            raise Unsupported('"'+reference['Name']+'" is declared in an unknown file')
        if reference['Api'] == self.api:
            return reference['Name']
        if reference['Api'] not in self.imports:
            self.imports.append(reference['Api'])
        return reference['Api']+'Bindings.'+reference['Name']

    def Type (self, type):
        '''The comruntime type of a described type.'''
        if type['Kind'] == 'Native':
            if type['Name'] == 'VARIANT_BOOL':
                return 'comruntime.Bool'
            if type['Name'] == 'BSTR':
                return 'comruntime.String'
            if type['Name'] in SCALAR:
                return 'comruntime.Value('+SCALAR[type['Name']]+')'
        elif type['Kind'] == 'ApiRef' and not IsInterface(type):
            return 'comruntime.Of('+self.ClassName(type)+')' # a struct or enum
        elif type['Kind'] == 'PointerTo' and IsInterface(type['Child']):
            return 'comruntime.InterfaceOf('+self.ClassName(type['Child'])+')'
        elif type['Kind'] == 'SafeArray':
            if type['Child'] is None:
                raise Unsupported('SAFEARRAY, whose element type the .winmd does not keep')
            return 'comruntime.SafeArrayOf('+self.Type(type['Child'])+')'
        elif type['Kind'] == 'PointerTo' and type['Child']['Kind'] == 'SafeArray':
            return self.Type(type['Child']) # a SAFEARRAY is passed as a SAFEARRAY *
        elif type['Kind'] == 'Array':
            size = 'None'
            if type['Shape']:
                size = str(type['Shape']['Size'])
            return 'comruntime.FixedArray('+self.Type(type['Child'])+', '+size+')'
        raise Unsupported('unsupported type "'+Describe(type)+'"')

    def Param (self, param):
        '''How a method argument is passed, as a comruntime parameter.'''
        name = repr(PythonName(param['Name']))
        direction = 'In'
        if 'Out' in param['Attrs']:
            direction = 'Out'
            if 'In' in param['Attrs']:
                direction = 'InOut'

        type = param['Type']
        if type['Kind'] == 'PointerTo' and type['Child']['Kind'] == 'PointerTo' and type['Child']['Child']['Kind'] == 'Native':
            raise Unsupported('unsupported type "'+Describe(type)+'"') # such as int **

        if type['Kind'] == 'Array':
            # passed as a pointer to its first element, as in C
            if direction == 'Out' and not type['Shape']:
                raise Unsupported('[out] array "'+param['Name']+'" has no fixed size')
            if direction == 'InOut':
                raise Unsupported('[in,out] array "'+param['Name']+'"')
            return 'comruntime.'+direction+'Array('+name+', '+self.Type(type)+')'

        if type['Kind'] == 'PointerTo' and not IsInterface(type['Child']):
            # a pointer to where the value is, or is written to
            if direction == 'In':
                direction = 'InPointer'
            return 'comruntime.'+direction+'('+name+', '+self.Type(type['Child'])+')'

        if direction != 'In':
            raise Unsupported('[out] argument "'+param['Name']+'" is not a pointer')
        return 'comruntime.In('+name+', '+self.Type(type)+')'

    def Method (self, interface, method):
        try:
            params = ''
            for param in method['Params']:
                params += ',\n        '+self.Param(param)
            return "    comruntime.Method('"+method['Name']+"'"+params+'),\n'
        except Unsupported as error:
            # keeps its slot, so that the other methods are unaffected
            self.unsupported.append(interface['Name']+'.'+method['Name']+': '+str(error))
            return "    comruntime.Unsupported('"+method['Name']+"', "+repr(str(error))+'),\n'

    def Generate (self):
        declared = '' # the classes, which the signatures below can then refer to
        signatures = ''

        for type in self.types:
            if type['Kind'] != 'Enum':
                continue
            declared += '\n\nclass '+type['Name']+' (comruntime.Enum):\n'
            for value in type['Values']:
                declared += '    '+PythonName(value['Name'])+' = '+str(value['Value'])+'\n'
            for value in type['Values']: # also by their C name, as in C
                declared += PythonName(value['Name'])+' = '+type['Name']+'.'+PythonName(value['Name'])+'\n'

        for type in self.types:
            if type['Kind'] != 'Struct':
                continue
            declared += '\n\nclass '+type['Name']+' (comruntime.Struct):\n    pass\n'
            try:
                fields = ''
                for field in type['Fields']:
                    fields += '    ('+repr(PythonName(field['Name']))+', '+self.Type(field['Type'])+'),\n'
                signatures += '\n'+type['Name']+'.FIELDS = [\n'+fields+']\n'
            except Unsupported as error:
                self.unsupported.append(type['Name']+': '+str(error))
                signatures += '\n'+type['Name']+'.UNSUPPORTED = '+repr(str(error))+'\n'

        for type in self.types:
            if type['Kind'] != 'Com':
                continue
            base = self.ClassName(type['Interface'])
            declared += '\n\nclass '+type['Name']+' ('+base+'):\n'
            declared += "    IID = '"+type['Guid']+"'\n"

            signatures += '\ncomruntime.Define('+type['Name']+', [\n'
            for method in type['Methods']:
                signatures += self.Method(type, method)
            signatures += '])\n'

            signatures += '\n\ndef Implement'+type['Name']+' (handler):\n'
            signatures += "    '''Back a "+type['Name']+" with a Python object, whose methods get the [in]\n"
            signatures += "       arguments and return the [out] ones: the value when there is one, and a\n"
            signatures += "       tuple in declaration order when there are several.'''\n"
            signatures += '    return comruntime.Implementation(handler, '+type['Name']+')\n\n'

        for type in self.types:
            if type['Kind'] != 'ComClassID':
                continue
            signatures += '\n\ndef Create'+type['Name']+' (library, interface):\n'
            signatures += "    '''Activate the "+type['Name']+" class through the library's CoCreateInstance.'''\n"
            signatures += "    return comruntime.CreateInstance(library, '"+type['Guid']+"', interface)\n"

        out = '# Generated from the '+self.api+' namespace of a .winmd file. Do not edit.\n'
        out += 'import ctypes\n'
        out += 'import comruntime\n'
        for api in self.imports:
            out += 'import '+api+'Bindings\n'
        return out+declared+'\n\n# signatures, which can refer to any class above\n'+signatures


def Describe (type):
    '''A type tree as C-like text, for messages.'''
    if type['Kind'] == 'PointerTo':
        if type['Child']['Kind'] == 'PointerTo':
            return Describe(type['Child'])+'*'
        return Describe(type['Child'])+' *'
    if type['Kind'] == 'Array':
        return Describe(type['Child'])+'[]'
    return type.get('Name', type['Kind'])


def CheckSources (winmd):
    '''Warns when an IDL file changed since the .winmd was generated from it,
       which only a Windows build can fix. The hashes are the ones that
       Winmd/WinmdRecordSources.py records, of the text with LF line endings.'''
    record = winmd+'.sha256'
    if not os.path.exists(record):
        print('winmd2py: warning: '+record+' is missing, so whether '+os.path.basename(winmd)+' is up to date is unknown')
        return
    folder = os.path.dirname(os.path.abspath(winmd))
    with open(record, 'r') as f:
        for line in f.read().splitlines():
            expected, name = line.split('  ', 1)
            with open(os.path.join(folder, name), 'rb') as source:
                actual = hashlib.sha256(source.read().replace(b'\r\n', b'\n')).hexdigest()
            if actual != expected:
                print('winmd2py: warning: '+name+' changed since '+os.path.basename(winmd)+' was generated from it, '+
                      'so the bindings describe the version before. A Windows build with the .NET SDK regenerates it.')


if __name__ == '__main__':
    if len(sys.argv) != 4:
        raise Exception('usage: winmd2py.py <input.winmd> <namespace> <output.py>')

    CheckSources(sys.argv[1])

    bindings = Bindings(sys.argv[1], sys.argv[2])
    with open(sys.argv[3], 'w') as f:
        f.write(bindings.Generate())

    # a method with an argument that cannot be passed yet raises
    # NotImplementedError when called, and E_NOTIMPL when implemented in Python
    for reason in bindings.unsupported:
        print('winmd2py: left out '+reason)
