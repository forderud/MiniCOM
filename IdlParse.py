# Simple Microsoft IDL parser.
# Generates cross-platform compatible C++ headers from Microsoft IDL files.

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
    coclasses = {}

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
                coclasses[coclass] = FindUuidString(substr)
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
    return modified, interfaces, coclasses


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


def DescribeType (type_text, known):
    '''Describe a C++ type as a tree, the way win32json does, so that a binding
       generator does not have to parse C++ declarators itself.'''
    if type_text.endswith('*'):
        return {'Kind': 'PointerTo', 'Child': DescribeType(type_text[:-1].strip(), known)}

    # "SAFEARRAY(T)" is declared as a SAFEARRAY pointer in the header. win32json
    # has no node for it, since it drops the element type, which is kept here.
    safearray = re.match('SAFEARRAY\\s*\\((.*)\\)$', type_text)
    if safearray:
        return {'Kind': 'SafeArray', 'Child': DescribeType(safearray.group(1).strip(), known)}

    if type_text == 'IUnknown':
        return {'Kind': 'ApiRef', 'Name': type_text, 'TargetKind': 'Com'}
    if type_text in known:
        # an interface ("Com"), or a struct or enum ("Default"), declared in
        # this IDL file or one it imports
        described = {'Kind': 'ApiRef', 'Name': type_text, 'TargetKind': known[type_text]['TargetKind']}
        if known[type_text]['Api']:
            described['Api'] = known[type_text]['Api'] # the IDL file that declares it
        return described
    return {'Kind': 'Native', 'Name': type_text}


def DescribeDeclaration (declaration, known):
    '''Take a "const IFoo ** obj" declaration apart, into its name and type.'''
    declaration = re.sub('\\bconst\\b', '', MASK_PATTERN.sub(' ', declaration)).strip()

    # "float res[3]" declares an array of a fixed, or with "[]" unknown, size
    array = re.search('\\[([0-9]*)\\]\\s*$', declaration)
    if array:
        declaration = declaration[:array.start()].strip()

    # the name is the identifier at the end, and the type is what comes before it
    name = ''
    type_text = declaration
    match = re.search('([a-zA-Z0-9_]+)$', declaration)
    if match:
        name = match.group(1)
        type_text = declaration[:match.start()]
    type_text = re.sub('\\s+', ' ', type_text.strip())
    type_text = re.sub('\\s*\\*', '*', type_text)

    described = DescribeType(type_text, known)
    if array:
        shape = None # size unknown
        if array.group(1):
            shape = {'Size': int(array.group(1))}
        described = {'Kind': 'Array', 'Shape': shape, 'Child': described}
    return name, described


def DescribeParameter (attributes, declaration, known):
    '''Describe a method parameter, with its direction as win32json "Attrs".'''
    attrs = []
    for attribute in attributes:
        if attribute in ['in', 'out', 'retval']:
            attrs.append({'in': 'In', 'out': 'Out', 'retval': 'RetVal'}[attribute])
    if not attrs:
        attrs = ['In']

    name, described = DescribeDeclaration(declaration, known)
    return {'Name': name, 'Type': described, 'Attrs': attrs}


# 'typedef [...] struct Name {...} Name;' and the same for enums, with comments
# replaced by placeholders
TYPEDEF_PATTERN = re.compile('typedef\\s+(?:\\[[^\\]]*\\]\\s*)?(?:'+MASK_BEGIN+'\\d+'+MASK_END+'\\s*)*(struct|enum)\\s*[a-zA-Z0-9_]*\\s*{(.*?)}\\s*([a-zA-Z0-9_]+)\\s*;', re.DOTALL)


def DeclaredTypes (source, api):
    '''The types an IDL source declares, as {name: {"TargetKind", "Api"}}, where
       "TargetKind" is what win32json has on a reference to it: "Com" for an
       interface, and "Default" for a struct or enum. "Api" is the file that
       defines it, which is not known for an interface that is only forward
       declared. Expects strings and comments to be placeholders.'''
    types = {}
    for name in re.findall('\\binterface\\s+([a-zA-Z0-9_]+)\\s*;', source):
        types[name] = {'TargetKind': 'Com', 'Api': None}
    for name, base, uuid, methods in CollectInterfaces(source):
        types[name] = {'TargetKind': 'Com', 'Api': api}
    for match in TYPEDEF_PATTERN.finditer(source):
        types[match.group(3)] = {'TargetKind': 'Default', 'Api': api}
    return types


def ApiName (idl_file):
    '''What a reference's "Api" field names: the IDL file that declares it.'''
    return os.path.splitext(os.path.basename(idl_file))[0]


def ImportedTypes (idl_file, seen):
    '''The types that the IDL files this one imports define, directly or
       through another import, as DeclaredTypes has them. These are looked up
       beside the importing file; system ones such as oaidl.idl are not there,
       and skipped.'''
    with open(idl_file, 'r') as f:
        source = f.read()

    types = {}
    for imported in re.findall('import\\s*"([a-zA-Z0-9_\\.]+)"\\s*;', source):
        path = os.path.join(os.path.dirname(idl_file), imported)
        if path in seen or not os.path.exists(path):
            continue
        seen.append(path)

        with open(path, 'r') as f:
            masked = []
            imported_source = ExtractComments(ExtractStrings(f.read(), masked), masked)
        declared = DeclaredTypes(imported_source, ApiName(path))
        for name in declared:
            if declared[name]['Api']:
                types[name] = declared[name]
        types.update(ImportedTypes(path, seen))
    return types


def EnumValue (expression, values):
    '''The value of a C constant expression such as "1 << 3" or "OTHER + 1", in
       which 'values' are the enumerators declared before it.'''
    expression = re.sub('(\\d)[uUlL]+\\b', '\\1', expression) # 1u, 1L
    for name in re.findall('[a-zA-Z_][a-zA-Z0-9_]*', expression):
        if name.startswith('0x') or name.startswith('0X'):
            continue
        if name not in values:
            raise Exception('unknown enum value "'+name+'" in "'+expression+'"')
        expression = re.sub('\\b'+name+'\\b', str(values[name]), expression)
    if not re.match('^[0-9a-fA-FxX\\s()+\\-*/%<>|&^~]*$', expression):
        raise Exception('unsupported enum expression "'+expression+'"')
    return int(eval(expression.replace('/', '//'), {'__builtins__': {}}))


def ParseTypedefs (source, known):
    '''Describe the structs and enums, win32json style, from the source after
       ParseAttributes, where comments are placeholders.'''
    types = []
    values = {} # every enumerator so far, which later ones can refer to
    for match in TYPEDEF_PATTERN.finditer(source):
        kind, body, name = match.group(1), match.group(2), match.group(3)
        body = MASK_PATTERN.sub(' ', body)

        if kind == 'enum':
            described = []
            value = 0
            for item in body.split(','): # not SplitParameters, which reads "<<" as brackets
                if not item.strip():
                    continue
                parts = item.split('=', 1)
                if len(parts) == 2:
                    value = EnumValue(parts[1].strip(), values)
                values[parts[0].strip()] = value
                described.append({'Name': parts[0].strip(), 'Value': value})
                value += 1
            types.append({'Name': name, 'Kind': 'Enum', 'Values': described})
        else:
            fields = []
            for declaration in body.split(';'):
                if not declaration.strip():
                    continue
                # "int a, b" declares two fields of the same type
                names = SplitParameters(declaration)
                field, described = DescribeDeclaration(names[0], known)
                fields.append({'Name': field, 'Type': described})
                type_text = names[0][:names[0].rfind(field)]
                for other in names[1:]:
                    field, described = DescribeDeclaration(type_text+' '+other, known)
                    fields.append({'Name': field, 'Type': described})
            types.append({'Name': name, 'Kind': 'Struct', 'Fields': fields})
    return types


def GenerateMetadata (definitions, coclasses, typedefs, known):
    '''Describe the interfaces as JSON, next to the generated header.

    MIDL writes a type library beside the header, which is what comtypes reads
    to build Python bindings on Windows. There is no type library here, so a
    binding generator would otherwise have to parse the IDL a second time. This
    writes what it needs instead, laid out like win32json (the JSON translation
    of Microsoft's win32metadata): every interface with its base, its IID and
    its own methods in declaration order, every struct with its fields and
    every enum with its values, and every coclass with its CLSID.
    '''
    types = []
    for name, base, uuid, methods in definitions:
        described = []
        for method, params in methods:
            described_params = []
            for attributes, declaration in params:
                described_params.append(DescribeParameter(attributes, declaration, known))
            described.append({'Name': method, 'Params': described_params})

        types.append({
            'Name': name,
            'Kind': 'Com',
            'Guid': uuid,
            'Interface': DescribeType(base, known),
            'Methods': described,
        })

    types += typedefs

    for name in coclasses:
        types.append({'Name': name, 'Kind': 'ComClassID', 'Guid': coclasses[name]})

    return json.dumps({'Types': types}, indent=2)+'\n'


def ParseIdlFile (idl_file, h_file, c_file, json_file):
    with open(idl_file, 'r') as f:
        source = f.read()

    # parse IDL file
    masked = []
    source = RemoveMidPragmas(source)
    source = ExtractStrings(source, masked)
    source = ExtractComments(source, masked)
    definitions = CollectInterfaces(source)

    # the types that arguments can refer to by name: those declared here, then
    # those defined in an imported file, unless also defined here
    known = DeclaredTypes(source, ApiName(idl_file))
    imported = ImportedTypes(idl_file, [])
    for name in imported:
        if name not in known or not known[name]['Api']:
            known[name] = imported[name]

    source, interfaces, coclasses = ParseAttributes(source)
    typedefs = ParseTypedefs(source, known)
    source = ParseInterfaces(source)
    source = ParseSafeArray(source)
    source = ReplaceComments(source, masked)
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
        f.write(GenerateMetadata(definitions, coclasses, typedefs, known))


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
