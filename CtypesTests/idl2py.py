# Generates ctypes bindings from the JSON description that IdlParse.py writes.
#
# Nothing here parses IDL. The interfaces, structs and enums come from the
# JSON, the same way comtypes reads the type library that MIDL writes on
# Windows. The bindings only declare them: an enum class, a struct class with
# its fields, and an interface class with the parameters of each method.
# Moving values between Python and C++ is left to comruntime.py, which is the
# same for every IDL file. What is left here is a table of base type names,
# which is per language.

import json
import keyword
import os
import sys

# mirrors the typedefs in NonWindows.hpp. A ctypes type has the size that the
# C compiler gives the C type it is named after on the platform it runs on,
# which is also what the C++ side of the bindings is compiled with. So "long",
# and BOOL, which NonWindows.hpp declares as "long", are c_long: 32bit on
# Windows, and 64bit on 64bit Linux and macOS.
SCALAR = {
    'char':               'ctypes.c_char',
    'wchar_t':            'ctypes.c_wchar',
    'short':              'ctypes.c_short',
    'int':                'ctypes.c_int',
    'long':               'ctypes.c_long',
    'long long':          'ctypes.c_longlong',
    '__int64':            'ctypes.c_longlong',
    'unsigned char':      'ctypes.c_ubyte',
    'unsigned short':     'ctypes.c_ushort',
    'unsigned int':       'ctypes.c_uint',
    'unsigned long':      'ctypes.c_ulong',
    'unsigned long long': 'ctypes.c_ulonglong',
    'unsigned __int64':   'ctypes.c_ulonglong',
    'float':              'ctypes.c_float',
    'double':             'ctypes.c_double',
    'BYTE':               'ctypes.c_ubyte',
    'USHORT':             'ctypes.c_ushort',
    'UINT':               'ctypes.c_uint',
    'DWORD':              'ctypes.c_uint',
    'ULONG':              'ctypes.c_uint',
    'LONG':               'ctypes.c_int',
    'BOOL':               'ctypes.c_long',
    'ULONGLONG':          'ctypes.c_ulonglong',
    'HRESULT':            'ctypes.c_int32',
    'HWND':               'ctypes.c_void_p',
}


class Unsupported (Exception):
    '''A type that the bindings cannot pass yet.'''


def PythonName (name):
    '''A name that is safe as a Python identifier in the generated code.'''
    if keyword.iskeyword(name) or name in ['self', 'ctypes', 'comruntime']:
        return name+'_'
    return name


def IsInterface (type):
    return type['Kind'] == 'ApiRef' and type['TargetKind'] == 'Com'


class Bindings:
    '''Generates the bindings for one JSON file. A type declared in another IDL
       file is taken from the bindings for that file, which are imported as
       <file>Bindings.'''

    def __init__ (self, json_file):
        self.api = os.path.splitext(os.path.basename(json_file))[0]
        with open(json_file, 'r') as f:
            self.types = json.load(f)['Types']
        self.imports = []     # other files whose bindings are referred to
        self.unsupported = [] # what is left out, with the reason

    def ClassName (self, reference):
        '''The Python class for a type that the JSON refers to by name.'''
        if 'Api' not in reference:
            if IsInterface(reference):
                return 'comruntime.Interface' # IUnknown, or declared in an unknown file
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
            return 'comruntime.SafeArrayOf('+self.Type(type['Child'])+')'
        elif type['Kind'] == 'Array':
            size = 'None'
            if type['Shape']:
                size = str(type['Shape']['Size'])
            return 'comruntime.FixedArray('+self.Type(type['Child'])+', '+size+')'
        raise Unsupported('unsupported type "'+type.get('Name', type['Kind'])+'"')

    def Param (self, param):
        '''How a method argument is passed, as a comruntime parameter.'''
        name = repr(PythonName(param['Name']))
        direction = 'In'
        if 'Out' in param['Attrs']:
            direction = 'Out'
            if 'In' in param['Attrs']:
                direction = 'InOut'

        type = param['Type']
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

        out = '# Generated from the IdlParse.py JSON description. Do not edit.\n'
        out += 'import ctypes\n'
        out += 'import comruntime\n'
        for api in self.imports:
            out += 'import '+api+'Bindings\n'
        return out+declared+'\n\n# signatures, which can refer to any class above\n'+signatures


if __name__ == '__main__':
    if len(sys.argv) != 3:
        raise Exception('usage: idl2py.py <input.json> <output.py>')

    bindings = Bindings(sys.argv[1])
    with open(sys.argv[2], 'w') as f:
        f.write(bindings.Generate())

    # a method with an argument that cannot be passed yet raises
    # NotImplementedError when called, and E_NOTIMPL when implemented in Python
    for reason in bindings.unsupported:
        print('idl2py: left out '+reason)
