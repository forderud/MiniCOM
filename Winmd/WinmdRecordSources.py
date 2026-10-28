# Records the IDL files a .winmd file was generated from, so that a build that
# only reads it can tell when it is out of date.
#
# The IDL files are recorded in <file>.winmd.sha256, in the format of
# "shasum -a 256", hashed with LF line endings so that a checkout with CRLF
# ones gives the same hash.

import hashlib
import os
import sys


def SourceHash (path):
    '''The hash of a text file, the same whichever line endings it is checked out with.'''
    with open(path, 'rb') as f:
        return hashlib.sha256(f.read().replace(b'\r\n', b'\n')).hexdigest()


def RecordSources (path, sources):
    '''Writes <path>.sha256, with each source relative to the .winmd.'''
    with open(path+'.sha256', 'w', newline='\n') as f:
        for source in sources:
            name = os.path.relpath(source, os.path.dirname(os.path.abspath(path))).replace(os.sep, '/')
            f.write(SourceHash(source)+'  '+name+'\n')


if __name__ == '__main__':
    if len(sys.argv) < 3:
        raise Exception('usage: WinmdRecordSources.py <file.winmd> <file.idl>...')
    RecordSources(sys.argv[1], sys.argv[2:])
