# Makes a .winmd file the same for the same interfaces, so that it can be
# committed without changing on every build.
#
# The generator stamps each file with the time it was written, in the PE
# header, and with a random module version id (MVID). The time is zeroed, and
# the MVID is made from a hash of the rest of the file instead, so that it
# still changes when the interfaces do, which is what it is for.
#
# This script can be removed once https://github.com/microsoft/win32metadata/pull/2396
# is merged into win32metadata and released in Microsoft.Windows.WinmdGenerator,
# since that PR makes the generator write the file deterministically itself.

import hashlib
import struct
import sys
import uuid

import dnfile


def Normalize (path):
    pe = dnfile.dnPE(path)
    stamp = pe.FILE_HEADER.get_file_offset() + 4 # TimeDateStamp
    mvid = bytes(pe.net.mdtables.Module[0].Mvid.raw_data)
    pe.close()

    with open(path, 'rb') as f:
        data = bytearray(f.read())

    where = data.find(mvid)
    if where < 0 or data.find(mvid, where+1) >= 0:
        raise Exception('the MVID of '+path+' is not found exactly once')

    data[stamp:stamp+4] = struct.pack('<I', 0)
    data[where:where+16] = bytes(16)
    data[where:where+16] = uuid.UUID(bytes=hashlib.sha256(data).digest()[:16]).bytes_le

    with open(path, 'wb') as f:
        f.write(data)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        raise Exception('usage: WinmdRemoveTimeStamps.py <file.winmd>')
    Normalize(sys.argv[1])
