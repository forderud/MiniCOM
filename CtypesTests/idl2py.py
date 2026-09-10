# Generates ctypes bindings from the JSON description that IdlParse.py writes.
#
# Nothing here parses IDL or knows the vtable layout. Both come from the JSON,
# the same way comtypes reads the type library that MIDL writes on Windows.
# What is left is a table of base type names, which is per language.

import json
import sys

import comruntime

# mirrors the typedefs in NonWindows.hpp
SCALAR = {
    'int':            'ctypes.c_int',
    'short':          'ctypes.c_short',
    'long':           'ctypes.c_long',
    'float':          'ctypes.c_float',
    'double':         'ctypes.c_double',
    'unsigned char':  'ctypes.c_ubyte',
    'unsigned short': 'ctypes.c_ushort',
    'unsigned int':   'ctypes.c_uint',
    'BYTE':           'ctypes.c_ubyte',
    'USHORT':         'ctypes.c_ushort',
    'UINT':           'ctypes.c_uint',
    'DWORD':          'ctypes.c_uint',
    'LONG':           'ctypes.c_int',
    'BOOL':           'ctypes.c_long',
    'ULONGLONG':      'ctypes.c_ulonglong',
    'VARIANT_BOOL':   'ctypes.c_short',
    'HRESULT':        'ctypes.c_int32',
    'BSTR':           'ctypes.c_wchar_p',
}


def PointerLevels (argument):
    '''An interface is always reached through a pointer, absorbed by c_void_p.'''
    return argument['pointer']-1 if argument['kind'] == 'interface' else argument['pointer']


def CtypesFor (argument, levels=None):
    '''ctypes type of a described argument, or of what it points at.'''
    if argument['kind'] == 'interface':
        text = 'ctypes.c_void_p'
    elif argument['type'] in SCALAR:
        text = SCALAR[argument['type']]
    else:
        raise Exception('unsupported type "'+argument['type']+'"')

    for _ in range(PointerLevels(argument) if levels is None else levels):
        text = 'ctypes.POINTER('+text+')'
    return text


def SplitRetval (method):
    '''An [out,retval] argument is returned rather than passed, the way the
       generated C++ wrappers and comtypes both do it.'''
    arguments = method['arguments']
    if arguments and 'retval' in arguments[-1]['direction']:
        return arguments[:-1], arguments[-1]
    return arguments, None


def GenerateInterface (interface):
    '''A class that calls through the vtable of an object the library handed out.'''
    out = '\n\nclass '+interface['name']+' (comruntime.Interface):\n'
    out += "    IID = '"+interface['iid']+"'\n"

    for method in interface['methods']:
        passed, retval = SplitRetval(method)
        types = [CtypesFor(a) for a in method['arguments']]
        args  = [a['name'] for a in passed]

        out += '\n    def '+method['name']+' (self'+''.join(', '+a for a in args)+'):\n'
        if retval:
            out += '        '+retval['name']+' = '+CtypesFor(retval, PointerLevels(retval)-1)+'()\n'
            args.append('ctypes.byref('+retval['name']+')')

        out += '        hr = comruntime.Call(self.pointer, '+str(method['slot'])+', '
        out += '['+', '.join(types)+']'+''.join(', '+a for a in args)+')\n'
        out += "        comruntime.Check(hr, '"+interface['name']+'.'+method['name']+"')\n"

        if retval and retval['kind'] == 'interface':
            out += '        return '+retval['type']+'('+retval['name']+')\n'
        elif retval:
            out += '        return '+retval['name']+'.value\n'
    return out


def GenerateImplementation (interface, iids):
    '''A function that backs the interface with a Python object, for the library
       to call into.'''
    out = '\n\ndef Implement'+interface['name']+' (handler):\n'
    out += "    '''Back a "+interface['name']+" with a Python object. A handler method gets the\n"
    out += "       [in] arguments, and what it returns becomes the [out,retval] one.'''\n"
    out += '    methods = [\n'

    for method in interface['methods']:
        passed, retval = SplitRetval(method)
        types = [CtypesFor(a) for a in method['arguments']]
        names = [a['name'] for a in method['arguments']]

        out += '        ctypes.CFUNCTYPE('+', '.join(['ctypes.c_int32', 'ctypes.c_void_p']+types)+')(\n'
        out += '            lambda this'+''.join(', '+n for n in names)+': comruntime.Returned(\n'
        out += '                '+(retval['name'] if retval else 'None')+', '
        out += 'handler.'+method['name']+'('+', '.join(a['name'] for a in passed)+'))),\n'

    out += '    ]\n'
    out += '    return comruntime.Implementation(handler, [\n'
    for iid in iids:
        out += "        '"+iid+"',\n"
    out += '    ], methods)\n'
    return out


def GenerateBindings (metadata):
    iid_of  = {i['name']: i['iid'] for i in metadata['interfaces']}
    base_of = {i['name']: i['base'] for i in metadata['interfaces']}

    def IidChain (name):
        '''The interface, everything it inherits, and IUnknown. What an object
           backed by a Python handler has to answer QueryInterface for.'''
        chain = []
        while name in iid_of:
            chain.append(iid_of[name])
            name = base_of[name]
        return chain + [comruntime.IID_IUnknown]

    out = '# Generated from the IdlParse.py JSON description. Do not edit.\n'
    out += 'import ctypes\n'
    out += 'import comruntime\n'

    for interface in metadata['interfaces']:
        out += GenerateInterface(interface)
        out += GenerateImplementation(interface, IidChain(interface['name']))

    for coclass in metadata['coclasses']:
        out += '\n\ndef Create'+coclass['name']+' (library, interface):\n'
        out += "    '''Activate the "+coclass['name']+" class through the library's CoCreateInstance.'''\n"
        out += "    return comruntime.CreateInstance(library, '"+coclass['clsid']+"', interface)\n"

    return out


if __name__ == '__main__':
    if len(sys.argv) != 3:
        raise Exception('usage: idl2py.py <input.json> <output.py>')

    with open(sys.argv[1], 'r') as f:
        metadata = json.load(f)

    with open(sys.argv[2], 'w') as f:
        f.write(GenerateBindings(metadata))
