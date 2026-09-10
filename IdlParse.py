# Simple Microsoft IDL parser.
# Generates cross-platform compatible C++ headers from Microsoft IDL files.

import hashlib
import json
import os
import re
import sys

VERBOSE = False #True


def RemoveMidPragmas (source):
    result = ''
    for line in source.splitlines():
        if 'midl_pragma' in line:
            continue # skip line
        result += line + '\n'
    return result

def ExtractComments (source, comments):
    '''Extract comments & replace them with a hash value'''

    def ReplaceFun (match):
        beginidx, endidx = match.regs[0]
        substr = source[beginidx:endidx]

        hash = hashlib.md5(substr.encode()).hexdigest()
        comments[hash] = substr
        return hash

    # pattern that detects multi-line "/*...*/" strings non-greedy
    pattern = re.compile('/\\*.*?\\*/', re.DOTALL)
    source = pattern.sub(ReplaceFun, source)

    # pattern that detects "//..." strings
    pattern = re.compile('//.*')
    source = pattern.sub(ReplaceFun, source)
    return source


def ExtractStrings (source, comments):
    '''Extract text strings & replace them with a hash value'''

    def ReplaceFun (match):
        beginidx, endidx = match.regs[0]
        substr = source[beginidx:endidx]

        hash = hashlib.md5(substr.encode()).hexdigest()
        comments[hash] = substr
        return hash

    # pattern that detects "..." strings that might contain escape characters
    pattern = re.compile('"([^"\\\\]|\\\\.)*"')
    source = pattern.sub(ReplaceFun, source)
    return source


def ReplaceComments (source, comments):
    '''Substitute hash values back with their original text strings'''
    for i in range(2): # two passes to account for nested comments
        for key in comments:
            source = source.replace(key, comments[key])
    return source


def FindUuidString (attributes):
    '''Return the uuid of an IDL '[...uuid(FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF)...]' attribute'''
    uuid = attributes[attributes.find('uuid(')+5:]
    return uuid[:uuid.find(')')]

def ParseUuidString (str):
    '''Return uuid string on {0xFFFFFFFF,0xFFFF,0xFFFF,{0xFF,0xFF,0xFF,0xFF,0xFF,0xFF,0xFF,0xFF}} form'''
    uuid = ''.join(FindUuidString(str).split('-'))
    uuid = '{0x'+uuid[:8]+',0x'+uuid[8:12]+',0x'+uuid[12:16]+',{0x'+uuid[16:18]+',0x'+uuid[18:20]+',0x'+uuid[20:22]+',0x'+uuid[22:24]+',0x'+uuid[24:26]+',0x'+uuid[26:28]+',0x'+uuid[28:30]+',0x'+uuid[30:32]+'}}'
    return uuid

def ParseAttributes (source):
    '''Parse IDL '[...]' attributes'''
    interfaces = []
    uuids = {}
    coclasses = {}

    def ReplaceFun (match):
        beginidx, endidx = match.regs[0]
        substr = source[beginidx:endidx]

        uuid_statement = ''
        if 'uuid(' in substr:
            # identify which struct/interface the uuid belongs to
            later_tokens = source[endidx:].split()
            if later_tokens[0] == 'interface':
                interface = later_tokens[1] # interface name
                # identify UUID
                uuid = ParseUuidString(substr)
                uuid_statement = '\nstatic constexpr GUID IID_'+interface+' = '+uuid+';\n'
                interfaces.append(interface)
                uuids[interface] = FindUuidString(substr)
            elif later_tokens[0] == 'coclass':
                coclass = later_tokens[1] # class name
                uuid = ParseUuidString(substr)
                uuid_statement = '\nstatic constexpr GUID CLSID_'+coclass+' = '+uuid+';\n'
                coclasses[coclass] = FindUuidString(substr)
        else:
            # preserve brackets in "float res[3]"-style arguments
            if substr[1:-1].isdigit():
                uuid_statement = substr
            else:
                # keep the argument direction in a comment, for GenerateMetadata
                # below. Stripped again before the header is written.
                directions = re.findall('\\b(in|out|retval)\\b', substr)
                if directions:
                    uuid_statement = '/*['+','.join(directions)+']*/'

        # sourround attributes in a comment
        result = ''
        if VERBOSE:
            result += '/*'+substr+'*/'
        else:
            result += ''
        result += uuid_statement
        return result

    # pattern to match multi-line '[...]' attributes non-greedy
    pattern = re.compile('\\[.*?\\]', re.DOTALL)
    modified = pattern.sub(ReplaceFun, source)
    return modified, interfaces, uuids, coclasses


def ParseInterfaces (source):
    '''Parse IDL interface statements.

    Also returns what the patterns below already match, as
    [(name, base, [(method, arglist), ...]), ...] in declaration order, so that
    the JSON description does not need to parse the result a second time.
    '''
    definitions = []

    def ReplaceFun1 (match):
        beginidx, endidx = match.regs[0]
        substr = source[beginidx:endidx]
        definitions.append((match.group(1), match.group(2), []))
        # rename 'interface' to 'struct'
        substr = substr.replace('interface', 'struct', 1)
        # add public inheritance
        return substr.replace(':', ': public', 1)

    # pattern to match 'interface ABC : IUnknown {'
    pattern = re.compile('interface\\s*([a-zA-Z0-9_]+?)\\s*:\\s*([a-zA-Z0-9_]+)\\s*{')
    source = pattern.sub(ReplaceFun1, source)

    def ReplaceFun2 (match):
        beginidx, endidx = match.regs[0]
        substr = source[beginidx:endidx]
        # attribute the method to the interface it is declared in
        index = source.count(': public', 0, beginidx)-1
        if index >= 0:
            definitions[index][2].append((match.group(1), match.group(2)))
        # add 'virtual' to signature
        substr = substr.replace('HRESULT', 'virtual HRESULT', 1)
        # add '= 0' after signature
        substr = substr[:-1]+' = 0;'
        return substr

    # pattern to match 'HRESULT Fun (....);' method signatures
    pattern2 = re.compile('HRESULT \\s*([a-zA-Z0-9_]+?)\\s*\\((.*?)\\)\\s*;', re.DOTALL)
    source = pattern2.sub(ReplaceFun2, source)
            
    # pattern to match 'coclass ABC {...};'
    interface_signature = '(\\[default\\])?\\s+interface\\s+[a-zA-Z0-9_]+;'
    pattern = re.compile('coclass\\s*[a-zA-Z0-9_]+?\\s*{(\\s*'+interface_signature+')*\\s*};')
    source = pattern.sub('', source) # remove matches
    
    def ReplaceFun3 (match):
        beginidx, endidx = match.regs[0]
        substr = source[beginidx:endidx]
        # rename 'interface' to 'struct'
        return substr.replace('interface', 'struct', 1)
    
    # pattern to match 'interface ABC;' (forward declarations)
    # note that this MUST be done after processing the 'coclass' definitions, since they contain interface listings
    pattern = re.compile('interface\\s*[a-zA-Z0-9_]+?\\s*;')
    source = pattern.sub(ReplaceFun3, source)
    
    return source, definitions


def ParseCppQuote (source):
    '''Parse cpp_quote("...") statements'''

    def ReplaceFun (match):
        beginidx, endidx = match.regs[0]
        substr = source[beginidx:endidx]
        substr = substr.replace('\\"', '"')
        return substr[11:-2]

    # pattern to match 'cpp_quote("...")'
    pattern = re.compile('cpp_quote\\(".*"\\)')
    source = pattern.sub(ReplaceFun, source)
    return source


def ParseSafeArray (source):
    '''Parse SAFEARRAY(T) statements'''

    def ReplaceFun (match):
        beginidx, endidx = match.regs[0]
        substr = source[beginidx:endidx]
        if VERBOSE:
            substr = substr.replace('(', '/*(')
            substr = substr.replace(')', ')*/')
            return substr+' * '
        else:
            return 'SAFEARRAY *'

    # pattern to match 'SAFEARRAY(byte)' and similar
    pattern = re.compile('SAFEARRAY\\([a-zA-Z0-9_\\s\\*]+?\\)') # matches a-z, A-Z, '_' & '*'
    # convert to SAFEARRAY pointer
    source = pattern.sub(ReplaceFun, source)
    return source


def ParseImport (source):
    '''Modify import "..." statements'''
    global last_import, found_lib
    last_import = 0
    found_lib = False
    
    def ReplaceFun (match):
        global last_import
        beginidx, endidx = match.regs[0]
        last_import = beginidx
        substr = source[beginidx:endidx]
        filename = substr[substr.find('"')+1:substr.rfind('"')]
        if filename.lower() in ['oaidl.idl', 'ocidl.idl']:
            return '' # remove import

        # change 'import "abc.idl";' to '#include "abc.h"'
        return '#include "'+filename[:filename.rfind('.')]+'.h"'
    
    pattern = re.compile('import \\s*"[a-zA-Z0-9_\\.]+?"\\s*;')
    source = pattern.sub(ReplaceFun, source)
    
    def RemoveFun (match):
        return ''
    
    # remove 'importlib("...");'
    pattern = re.compile('importlib\\("[a-zA-Z0-9_\\.]+?"\\);')
    source = pattern.sub(RemoveFun, source)
    
    def RemoveLibFun (match):
        global found_lib
        found_lib = True
        return ''
    
    # remove 'library XXX {'
    pattern = re.compile('library [a-zA-Z0-9_\\.]+\\s*{')
    source = pattern.sub(RemoveLibFun, source)
    
    if found_lib:
        # remove '};' at end of library scope
        idx = source.rfind('};')
        source = source[:idx] + source[idx+2:]
    
    last_import += source[last_import:].find('\n') # start of line after last import
    
    return source, last_import


def SplitParameters (arglist):
    '''Split a C++ parameter list on its top level commas.'''
    params = []
    depth = 0
    current = ''
    for ch in arglist:
        if ch in '([<':
            depth += 1
        elif ch in ')]>':
            depth -= 1
        if ch == ',' and depth == 0:
            params.append(current)
            current = ''
        else:
            current += ch
    params.append(current)
    return [p.strip() for p in params if p.strip()]


def DescribeArgument (argument, interfaces):
    '''Take a "/*[out]*/ IFoo ** obj" declaration apart, so that a binding
       generator does not have to parse C++ declarators itself.'''
    directions = re.findall('/\\*\\[(.*?)\\]\\*/', argument)
    declaration = re.sub('/\\*\\[.*?\\]\\*/', '', argument).replace('const', '').strip()

    # "float res[3]" declares an array of a fixed, or with "[]" unknown, size
    array = re.search('\\[([0-9]*)\\]\\s*$', declaration)
    if array:
        declaration = declaration[:array.start()].strip()

    name = re.search('([a-zA-Z0-9_]+)$', declaration)
    type_text = declaration[:name.start()] if name else declaration
    type_name = type_text.replace('*', '').strip()

    described = {
        'name': name.group(1) if name else '',
        'type': type_name,
        'pointer': type_text.count('*'),
        'kind': 'interface' if type_name in interfaces or type_name == 'IUnknown' else 'scalar',
        'direction': directions[0].split(',') if directions else ['in'],
    }
    if array:
        described['array'] = int(array.group(1)) if array.group(1) else 0
    return described


def GenerateMetadata (definitions, uuids, coclasses):
    '''Describe the interfaces as JSON, next to the generated header.

    MIDL writes a type library beside the header, which is what comtypes reads
    to build Python bindings on Windows. There is no type library here, so a
    binding generator would otherwise have to parse the IDL a second time and
    hard-code the vtable slot numbers. This writes what it needs instead: the
    inheritance chain, the interface IDs, and every method at the slot the
    generated interface puts it in.
    '''
    methods_of = {name: methods for name, _, methods in definitions}
    base_of    = {name: base for name, base, _ in definitions}

    def Inherits (name):
        '''Whether the whole inheritance chain is declared in this file.'''
        base = base_of.get(name)
        if not base or base == 'IUnknown':
            return base == 'IUnknown'
        return Inherits(base)

    def AllMethods (name):
        '''Own methods, preceded by every inherited one, in vtable order.'''
        base = base_of.get(name)
        inherited = AllMethods(base) if base != 'IUnknown' else []
        return inherited+methods_of[name]

    described = []
    for name, base, _ in definitions:
        if not Inherits(name):
            continue # slots unknown, since a base interface is declared elsewhere

        methods = []
        # QueryInterface, AddRef and Release occupy the first three slots
        for slot, (method, arglist) in enumerate(AllMethods(name), start=3):
            methods.append({
                'name': method,
                'slot': slot,
                'returns': 'HRESULT',
                'arguments': [DescribeArgument(a, base_of) for a in SplitParameters(arglist)],
            })
        described.append({
            'name': name,
            'iid': uuids.get(name, ''),
            'base': base,
            'methods': methods,
        })

    classes = [{'name': name, 'clsid': clsid} for name, clsid in coclasses.items()]
    return json.dumps({'interfaces': described, 'coclasses': classes}, indent=2)+'\n'


def ParseIdlFile (idl_file, h_file, c_file, json_file):
    with open(idl_file, 'r') as f:
        source = f.read()

    # parse IDL file
    comments = {}
    source = RemoveMidPragmas(source)
    source = ExtractStrings(source, comments)
    source = ExtractComments(source, comments)
    source, interfaces, uuids, coclasses = ParseAttributes(source)
    source, definitions = ParseInterfaces(source)
    source = ParseSafeArray(source)
    source = re.sub('/\\*\\[.*?\\]\\*/', '', source) # drop the direction markers again
    source = ReplaceComments(source, comments)
    source, last_import = ParseImport(source)
    source = ParseCppQuote(source)
        
    #print(source)
    with open(h_file, 'w') as f:
        f.write('#pragma once\n')
        f.write(source[:last_import]+'\n')
        f.write('extern "C" {\n')
        f.write(source[last_import:]+'\n')
        f.write('} //extern "C"\n')
        for interface in interfaces:
            f.write('DEFINE_UUIDOF('+interface+')\n')
    
    with open(c_file, 'w') as f:
        f.write('#include "'+h_file+'"\n')

    with open(json_file, 'w') as f:
        f.write(GenerateMetadata(definitions, uuids, coclasses))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise Exception('No IDL file arguments provided')

    # load files
    for filepath in sys.argv[1:]:
        #print('Parsing '+filepath)
        assert(filepath[-4:] == '.idl')
        
        # write generated headers to current dir.
        path, filename = os.path.split(filepath)
        ParseIdlFile(filepath, filename[:-4]+'.h', filename[:-4]+'_i.c', filename[:-4]+'.json')
