# Simple Microsoft IDL parser.
# Generates cross-platform compatible C++ headers from Microsoft IDL files.

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

#: Placeholder wrapper for masked text. Uses control characters so a placeholder
#: can never be produced by, or matched as part of, an IDL identifier.
MASK_BEGIN = '\x01'
MASK_END   = '\x02'
MASK_PATTERN = re.compile(MASK_BEGIN + r'(\d+)' + MASK_END)


def Mask (substr, masked):
    '''Stash substr and return a placeholder that no later pattern can match'''
    masked.append(substr)
    return MASK_BEGIN + str(len(masked) - 1) + MASK_END


def ExtractComments (source, masked):
    '''Replace comments with placeholders'''

    def ReplaceFun (match):
        return Mask(match.group(0), masked)

    # pattern that detects multi-line "/*...*/" strings non-greedy
    pattern = re.compile('/\\*.*?\\*/', re.DOTALL)
    source = pattern.sub(ReplaceFun, source)

    # pattern that detects "//..." strings
    pattern = re.compile('//.*')
    source = pattern.sub(ReplaceFun, source)
    return source


def ExtractStrings (source, masked):
    '''Replace text strings with placeholders'''

    def ReplaceFun (match):
        return Mask(match.group(0), masked)

    # pattern that detects "..." strings that might contain escape characters
    pattern = re.compile('"([^"\\\\]|\\\\.)*"')
    source = pattern.sub(ReplaceFun, source)
    return source


def ReplaceComments (source, masked):
    '''Substitute placeholders back with their original text'''
    # a masked comment can contain a placeholder for a string masked earlier, so
    # keep expanding until nothing is left rather than assuming a fixed depth
    while MASK_PATTERN.search(source):
        source = MASK_PATTERN.sub(lambda m: masked[int(m.group(1))], source)
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

    def ReplaceFun (match):
        substr = match.group(0)

        uuid_statement = ''
        if 'uuid(' in substr:
            # identify which struct/interface the uuid belongs to
            later_tokens = source[match.end():].split()
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
        substr = match.group(0)
        # rename 'interface' to 'struct'
        substr = substr.replace('interface', 'struct', 1)
        # add public inheritance
        return substr.replace(':', ': public', 1)

    # pattern to match 'interface ABC : IUnknown {'
    pattern = re.compile('interface\\s*[a-zA-Z0-9_]+?\\s*:\\s*[a-zA-Z0-9_]+\\s*{')
    source = pattern.sub(ReplaceFun1, source)

    def ReplaceFun2 (match):
        substr = match.group(0)
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
        substr = match.group(0)
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
        return match.group(1).replace('\\"', '"')

    # pattern to match 'cpp_quote("...")', capturing the quoted text. The body may
    # contain escaped quotes, so match those explicitly rather than stopping at the
    # first '"' -- and so that two statements on one line stay separate.
    pattern = re.compile('cpp_quote\\("((?:[^"\\\\]|\\\\.)*)"\\)', re.DOTALL)
    source = pattern.sub(ReplaceFun, source)
    return source


def ParseSafeArray (source):
    '''Parse SAFEARRAY(T) statements'''

    def ReplaceFun (match):
        substr = match.group(0)
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
    state = {'last_import': 0}

    def ReplaceFun (match):
        state['last_import'] = match.start()
        substr = match.group(0)
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
    
    # remove 'library XXX {'
    pattern = re.compile('library [a-zA-Z0-9_\\.]+\\s*{')
    match = pattern.search(source)
    if match:
        # find the '}' that closes the library scope, which is not
        # necessarily the last '};' in the file
        depth = 1
        idx = match.end()
        while depth > 0:
            if source[idx] == '{':
                depth += 1
            elif source[idx] == '}':
                depth -= 1
            idx += 1
        close = idx-1

        # also remove the ';' after the '}'
        if source[idx:idx+1] == ';':
            idx += 1

        source = source[:match.start()] + source[match.end():close] + source[idx:]
    
    last_import = state['last_import']
    last_import += source[last_import:].find('\n') # start of line after last import

    return source, last_import


def SplitParameters (arglist):
    '''Split a parameter or attribute list on its top level commas'''
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

    result = []
    for param in params:
        if param.strip():
            result.append(param.strip())
    return result


def CollectInterfaces (source):
    '''Return the interfaces an IDL source defines, in declaration order, as
       [(name, base, uuid, [(method, [(attributes, declaration), ...]), ...]), ...]

       'attributes' are those of a parameter, such as ['in'] or ['out', 'retval'],
       and 'declaration' is the rest of it, such as "SAFEARRAY(BYTE) * data".
       Expects strings and comments to have been replaced with placeholders.
       Leaves the source unchanged, so that the header does not depend on it.'''
    interfaces = []

    # pattern to match '[...] interface ABC : IUnknown {...}'
    pattern = re.compile('(\\[[^\\]]*\\])?\\s*interface\\s+([a-zA-Z0-9_]+)\\s*:\\s*([a-zA-Z0-9_]+)\\s*{(.*?)}', re.DOTALL)
    # pattern to match 'HRESULT Fun (....);' method signatures
    method_pattern = re.compile('HRESULT\\s+([a-zA-Z0-9_]+)\\s*\\((.*?)\\)\\s*;', re.DOTALL)

    for match in pattern.finditer(source):
        uuid = ''
        if match.group(1) and 'uuid(' in match.group(1):
            uuid = FindUuidString(match.group(1))

        methods = []
        for method in method_pattern.finditer(match.group(4)):
            params = []
            for param in SplitParameters(method.group(2)):
                attributes = []
                attribute = re.match('\\[(.*?)\\]', param)
                if attribute:
                    attributes = SplitParameters(attribute.group(1))
                    param = param[attribute.end():]
                # drop comments, and the line breaks of a multi-line declaration
                declaration = ' '.join(MASK_PATTERN.sub(' ', param).split())
                params.append((attributes, declaration))
            methods.append((method.group(1), params))

        interfaces.append((match.group(2), match.group(3), uuid, methods))
    return interfaces


def GenerateForwardDeclarations (interfaces):
    '''Forward declare every interface, the way MIDL does.

    Lets an interface refer to one that is declared further down the file,
    without depending on the order they appear in.
    '''
    out = ''
    for name in interfaces:
        out += 'typedef struct '+name+' '+name+';\n'
    return out+'\n'


def GenerateCInterfaces (source, definitions):
    '''Add C style vtable definitions next to the C++ interface definitions.

    Matches the IUnknown/IUnknownVtbl pair in NonWindows.hpp: the C++ interface
    by default, and a vtable struct with an "lpVtbl" object when CINTERFACE is
    defined. Both describe the same vtable, so a client can pick either form.
    '''
    methods_of = {}
    base_of = {}
    for name, base, uuid, methods in definitions:
        methods_of[name] = methods
        base_of[name] = base

    def AllMethods (name):
        '''Own methods, preceded by every inherited one, in vtable order.'''
        inherited = []
        base = base_of.get(name)
        if base and base != 'IUnknown':
            inherited = AllMethods(base)
        return inherited+methods_of.get(name, [])

    def BaseIsVisible (name):
        base = base_of.get(name)
        if not base or base == 'IUnknown':
            return True
        return base in methods_of and BaseIsVisible(base)

    def FindStruct (source, name):
        '''Span of the "struct <name> : public ..." definition, including the trailing ";".'''
        begin = re.search('struct\\s*'+name+'\\s*: public\\s*[a-zA-Z0-9_]+\\s*{', source)
        depth = 1
        idx = begin.end()
        while depth > 0:
            if source[idx] == '{':
                depth += 1
            elif source[idx] == '}':
                depth -= 1
            idx += 1
        return begin.start(), source.index(';', idx)+1

    # back to front, so that the offsets of earlier interfaces still apply
    for name, base, uuid, methods in reversed(definitions):
        if not BaseIsVisible(name):
            # skip interface that inherits unknown interface
            continue

        begin, end = FindStruct(source, name)

        vtbl = ''
        vtbl += '#if !defined(CINTERFACE)\n'
        vtbl += source[begin:end]
        vtbl += '\n#else // defined(CINTERFACE)\n\n'
        vtbl += 'typedef struct '+name+'Vtbl {\n'
        vtbl += '    HRESULT (*QueryInterface)('+name+'* This, const GUID& iid, void** obj);\n'
        vtbl += '    ULONG   (*AddRef)('+name+'* This);\n'
        vtbl += '    ULONG   (*Release)('+name+'* This);\n'
        for method, params in AllMethods(name):
            vtbl += '    HRESULT (*'+method+')('+name+'* This'
            for attributes, declaration in params:
                vtbl += ', '+ParseSafeArray(declaration)
            vtbl += ');\n'
        vtbl += '} '+name+'Vtbl;\n\n'
        vtbl += 'struct '+name+' {\n'
        vtbl += '    struct '+name+'Vtbl* lpVtbl;\n'
        vtbl += '};\n'
        vtbl += '#endif\n'

        source = source[:begin]+vtbl+source[end:]

    return source


def ParseIdlFile (idl_file, h_file, c_file):
    with open(idl_file, 'r') as f:
        source = f.read()

    # parse IDL file
    masked = []
    source = RemoveMidPragmas(source)
    source = ExtractStrings(source, masked)
    source = ExtractComments(source, masked)
    definitions = CollectInterfaces(source)
    source, interfaces = ParseAttributes(source)
    source = ParseInterfaces(source)
    source = ParseSafeArray(source)
    source = GenerateCInterfaces(source, definitions)
    source = ReplaceComments(source, masked)
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
