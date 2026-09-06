# Simple Microsoft IDL parser.
# Generates cross-platform compatible C++ headers from Microsoft IDL files.

import hashlib
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


def ParseUuidString (str):
    '''Return uuid string on {0xFFFFFFFF,0xFFFF,0xFFFF,{0xFF,0xFF,0xFF,0xFF,0xFF,0xFF,0xFF,0xFF}} form'''
    # identify UUID
    uuid = str[str.find('uuid(')+5:]
    uuid = ''.join(uuid[:uuid.find(')')].split('-'))
    uuid = '{0x'+uuid[:8]+',0x'+uuid[8:12]+',0x'+uuid[12:16]+',{0x'+uuid[16:18]+',0x'+uuid[18:20]+',0x'+uuid[20:22]+',0x'+uuid[22:24]+',0x'+uuid[24:26]+',0x'+uuid[26:28]+',0x'+uuid[28:30]+',0x'+uuid[30:32]+'}}'
    return uuid

def ParseAttributes (source):
    '''Parse IDL '[...]' attributes'''
    interfaces = []

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
            elif later_tokens[0] == 'coclass':
                coclass = later_tokens[1] # class name
                uuid = ParseUuidString(substr)
                uuid_statement = '\nstatic constexpr GUID CLSID_'+coclass+' = '+uuid+';\n'
        else:
            # preserve brackets in "float res[3]"-style arguments
            if substr[1:-1].isdigit():
                uuid_statement = substr

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
    return modified, interfaces


def ParseInterfaces (source):
    '''Parse IDL interface statements'''

    def ReplaceFun1 (match):
        beginidx, endidx = match.regs[0]
        substr = source[beginidx:endidx]
        # rename 'interface' to 'struct'
        substr = substr.replace('interface', 'struct', 1)
        # add public inheritance
        return substr.replace(':', ': public', 1)

    # pattern to match 'interface ABC : IUnknown {'
    pattern = re.compile('interface\\s*[a-zA-Z0-9_]+?\\s*:\\s*[a-zA-Z0-9_]+\\s*{')
    source = pattern.sub(ReplaceFun1, source)

    def ReplaceFun2 (match):
        beginidx, endidx = match.regs[0]
        substr = source[beginidx:endidx]
        # add 'virtual' to signature
        substr = substr.replace('HRESULT', 'virtual HRESULT', 1)
        # add '= 0' after signature
        substr = substr[:-1]+' = 0;'
        return substr

    # pattern to match 'HRESULT Fun (....);' method signatures
    pattern2 = re.compile('HRESULT \\s*[a-zA-Z0-9_]+?\\s*\\(.*?\\)\\s*;', re.DOTALL)
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
    
    return source


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
    '''Split a C++ parameter list on top-level commas.'''
    params, depth, current = [], 0, ''
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
    if current.strip():
        params.append(current)
    return [p.strip() for p in params if p.strip()]


def ParseInterfaces2 (source):
    '''Return [(name, base, methods, begin, end), ...] for each interface struct.

    Reads the already-rewritten C++ text, where interfaces have become
    "struct IFoo : public IBar {" and methods "virtual HRESULT Fun (...) = 0;".
    "begin" and "end" delimit the whole struct, including the trailing ";".
    '''
    result = []
    header = re.compile('struct\\s+([a-zA-Z0-9_]+)\\s*:\\s*public\\s+([a-zA-Z0-9_]+)\\s*{')
    method = re.compile('virtual\\s+HRESULT\\s+([a-zA-Z0-9_]+)\\s*\\((.*?)\\)\\s*=\\s*0\\s*;', re.DOTALL)

    for match in header.finditer(source):
        # find the matching close brace
        depth, idx = 1, match.end()
        while depth > 0 and idx < len(source):
            if source[idx] == '{':
                depth += 1
            elif source[idx] == '}':
                depth -= 1
            idx += 1
        body = source[match.end():idx-1]
        while idx < len(source) and source[idx] != ';':
            idx += 1
        methods = [(m.group(1), m.group(2)) for m in method.finditer(body)]
        result.append((match.group(1), match.group(2), methods, match.start(), idx+1))
    return result


def GenerateCInterfaces (source):
    '''Add C-style vtable definitions next to the C++ interface definitions.

    Mirrors what MIDL generates: the C++ interface when compiling C++, and a
    vtable struct plus an "lpVtbl" object when CINTERFACE is defined. Both
    describe the same vtable, so a client can pick whichever form suits it.
    '''
    interfaces = ParseInterfaces2(source)
    methods_of = {name: methods for name, _, methods, _, _ in interfaces}
    base_of    = {name: base for name, base, _, _, _ in interfaces}

    def AllMethods (name):
        '''Own methods, preceded by every inherited one, in vtable order.'''
        base = base_of.get(name)
        inherited = AllMethods(base) if base and base != 'IUnknown' else []
        return inherited+methods_of.get(name, [])

    def BaseIsVisible (name):
        base = base_of.get(name)
        if not base or base == 'IUnknown':
            return True
        return base in methods_of and BaseIsVisible(base)

    # back to front, so that the offsets of earlier interfaces still apply
    for name, base, _, begin, end in reversed(interfaces):
        if not BaseIsVisible(name):
            # Inherits from an interface declared in an imported IDL file, so
            # the inherited methods are unknown and the vtable is incomplete.
            continue

        vtbl = ''
        vtbl += '#if defined(__cplusplus) && !defined(CINTERFACE)\n'
        vtbl += source[begin:end]
        vtbl += '\n#else /* C style interface */\n'
        vtbl += 'typedef struct '+name+'Vtbl {\n'
        vtbl += '    BEGIN_INTERFACE\n'
        vtbl += '    HRESULT (STDMETHODCALLTYPE *QueryInterface) ('+name+'* This, REFIID riid, void** ppvObject);\n'
        vtbl += '    ULONG (STDMETHODCALLTYPE *AddRef) ('+name+'* This);\n'
        vtbl += '    ULONG (STDMETHODCALLTYPE *Release) ('+name+'* This);\n'
        for method, arglist in AllMethods(name):
            params = SplitParameters(arglist)
            vtbl += '    HRESULT (STDMETHODCALLTYPE *'+method+') ('+name+'* This'
            vtbl += ''.join(', '+p for p in params)+');\n'
        vtbl += '    END_INTERFACE\n'
        vtbl += '} '+name+'Vtbl;\n\n'
        vtbl += 'struct '+name+' {\n'
        vtbl += '    CONST_VTBL struct '+name+'Vtbl* lpVtbl;\n'
        vtbl += '};\n'
        vtbl += '#endif\n'

        source = source[:begin]+vtbl+source[end:]

    return source


def GenerateForwardDeclarations (interfaces):
    '''Forward declare every interface, the way MIDL does.

    Lets an interface refer to one that is declared further down the file, and
    lets generated headers refer to each other, without depending on order.
    '''
    out = ''
    for name in interfaces:
        out += '#ifndef __'+name+'_FWD_DEFINED__\n'
        out += '#define __'+name+'_FWD_DEFINED__\n'
        out += 'typedef struct '+name+' '+name+';\n'
        out += '#endif\n\n'
    return out


def ParseIdlFile (idl_file, h_file, c_file):
    with open(idl_file, 'r') as f:
        source = f.read()

    # parse IDL file
    comments = {}
    source = RemoveMidPragmas(source)
    source = ExtractStrings(source, comments)
    source = ExtractComments(source, comments)
    source, interfaces = ParseAttributes(source)
    source = ParseInterfaces(source)
    source = ParseSafeArray(source)
    source = GenerateCInterfaces(source)
    source = ReplaceComments(source, comments)
    source, last_import = ParseImport(source)
    source = ParseCppQuote(source)
        
    #print(source)
    with open(h_file, 'w') as f:
        f.write('#pragma once\n')
        f.write(source[:last_import]+'\n')
        f.write('extern "C" {\n')
        f.write(GenerateForwardDeclarations(interfaces))
        f.write(source[last_import:]+'\n')
        f.write('} //extern "C"\n')
        for interface in interfaces:
            f.write('DEFINE_UUIDOF('+interface+')\n')
    
    with open(c_file, 'w') as f:
        f.write('#include "'+h_file+'"\n')


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise Exception('No IDL file arguments provided')

    # load files
    for filepath in sys.argv[1:]:
        #print('Parsing '+filepath)
        assert(filepath[-4:] == '.idl')
        
        # write generated headers to current dir.
        path, filename = os.path.split(filepath)
        ParseIdlFile(filepath, filename[:-4]+'.h', filename[:-4]+'_i.c')
