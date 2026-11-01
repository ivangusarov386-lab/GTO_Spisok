#!/usr/bin/env python3
"""Build a source-only MS-OVBA project; Excel compiles the readable VBA on open.

Development dependencies: oletools==0.60.2 (includes olefile).
No VBA compiler, cached executable code, or original example macros are retained.
"""
from pathlib import Path
import hashlib
import math
import re
import struct
import olefile
from oletools.olevba import VBA_Parser, decompress_stream

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / 'vendor/protocol/template-vbaProject.bin'
EXPECTED = '0ced1464b3677e98f5e3a8c5d80135e18dc98dca39299f1a8cfd2a00999fbf9f'
FREE, END, FAT, NONE = 0xffffffff, 0xfffffffe, 0xfffffffd, 0xffffffff


def compress_raw(data):
    # MS-OVBA 2.4.1: uncompressed chunks have exactly 4096 bytes.
    out = bytearray(b'\x01')
    for i in range(0, len(data), 4096):
        out += struct.pack('<H', 0x3fff) + data[i:i + 4096].ljust(4096, b' ')
    return bytes(out)


def write_cfb(streams):
    """CFB v3 writer with regular FAT, MiniFAT and balanced directory trees."""
    entries = [dict(name='Root Entry', kind=5, children=[])]
    storage = {'': 0}
    for path in streams:
        parent = ''
        for component in path.split('/')[:-1]:
            key = parent + '/' + component if parent else component
            if key not in storage:
                storage[key] = len(entries)
                entries[storage[parent]]['children'].append(len(entries))
                entries.append(dict(name=component, kind=1, children=[]))
            parent = key
        entries[storage[parent]]['children'].append(len(entries))
        entries.append(dict(name=path.split('/')[-1], kind=2, data=streams[path], children=[]))
    sectors, fat = [], []

    def allocate(data):
        if not data:
            return END
        start = len(sectors)
        chunks = [data[i:i + 512].ljust(512, b'\0') for i in range(0, len(data), 512)]
        sectors.extend(chunks)
        fat.extend(range(start + 1, start + len(chunks)))
        fat.append(END)
        return start

    mini, minifat = bytearray(), []
    for e in entries:
        e.update(left=NONE, right=NONE, child=NONE, color=1, start=END, size=0)
        if e['kind'] != 2:
            continue
        data = e['data']
        e['size'] = len(data)
        if len(data) >= 4096:
            e['start'] = allocate(data)
        elif data:
            start, count = len(minifat), math.ceil(len(data) / 64)
            e['start'] = start
            mini += data.ljust(count * 64, b'\0')
            minifat.extend(range(start + 1, start + count))
            minifat.append(END)
    entries[0]['start'], entries[0]['size'] = allocate(mini), len(mini)
    mini_start = allocate(b''.join(struct.pack('<I', n) for n in minifat).ljust(math.ceil(len(minifat) / 128) * 512, b'\xff'))
    mini_sectors = math.ceil(len(minifat) / 128)

    def tree(ids, depth=0, deepest=None):
        if not ids:
            return NONE
        if deepest is None:
            deepest = int(math.log2(len(ids)))
        mid = len(ids) // 2
        e = entries[ids[mid]]
        e['left'] = tree(ids[:mid], depth + 1, deepest)
        e['right'] = tree(ids[mid + 1:], depth + 1, deepest)
        e['color'] = 0 if depth == deepest and depth > 0 else 1
        return ids[mid]

    for e in entries:
        e['child'] = tree(sorted(e['children'], key=lambda i: (len(entries[i]['name']), entries[i]['name'].upper())))
    directory = bytearray()
    for e in entries:
        name = (e['name'] + '\0').encode('utf-16le')
        directory += name.ljust(64, b'\0')
        directory += struct.pack('<HBBIII', len(name), e['kind'], e['color'], e['left'], e['right'], e['child'])
        directory += bytes(16 + 4 + 8 + 8)  # CLSID, state, creation/modification times
        directory += struct.pack('<IQ', e['start'], e['size'])
    dir_start = allocate(directory)
    nfat = 1
    while math.ceil((len(sectors) + nfat) / 128) != nfat:
        nfat = math.ceil((len(sectors) + nfat) / 128)
    assert nfat <= 109
    fat_ids = list(range(len(sectors), len(sectors) + nfat))
    fat += [FAT] * nfat
    fat += [FREE] * (nfat * 128 - len(fat))
    fat_bytes = b''.join(struct.pack('<I', n) for n in fat)
    sectors.extend(fat_bytes[i:i + 512] for i in range(0, len(fat_bytes), 512))
    header = bytearray.fromhex('d0cf11e0a1b11ae1') + bytes(16)
    header += struct.pack('<HHHHH', 0x3e, 3, 0xfffe, 9, 6) + bytes(6)
    header += struct.pack('<IIIIIIIII', 0, nfat, dir_start, 0, 4096, mini_start, mini_sectors, END, 0)
    header += b''.join(struct.pack('<I', n) for n in fat_ids + [FREE] * (109 - nfat))
    assert len(header) == 512
    return bytes(header) + b''.join(sectors)


def build():
    assert hashlib.sha256(TEMPLATE.read_bytes()).hexdigest() == EXPECTED
    template = olefile.OleFileIO(TEMPLATE)
    keep = ['ThisWorkbook', 'Module1']
    directory = decompress_stream(template.openstream('VBA/dir').read())
    module_start = directory.index(b'\x19\0')
    prefix = bytearray(directory[:module_start])
    record = prefix.index(b'\x0f\0\x02\0\0\0')
    prefix[record + 6:record + 8] = struct.pack('<H', len(keep))
    modules = []
    pos = module_start
    while directory[pos:pos + 2] == b'\x19\0':
        size = struct.unpack_from('<I', directory, pos + 2)[0]
        name = directory[pos + 6:pos + 6 + size].decode('cp1252')
        end = directory.index(b'\x2b\0\0\0\0\0', pos) + 6
        if name in keep:
            module = bytearray(directory[pos:end])
            offset = module.index(b'\x31\0\x04\0\0\0')
            module[offset + 6:offset + 10] = bytes(4)  # source begins at stream offset zero
            modules.append(module)
        pos = end
    project = template.openstream('PROJECT').read().decode('cp1252')
    project = re.sub(r'^Document=(?!ThisWorkbook/)[^\r\n]*\r?\n', '', project, flags=re.M)
    project = re.sub(r'^(?:Sheet1|Sheet2|ThisWorkbook1)=[^\r\n]*\r?\n', '', project, flags=re.M)
    streams = {
        'PROJECT': project.encode('cp1252'),
        'PROJECTwm': b''.join(n.encode('cp1252') + b'\0' + (n + '\0').encode('utf-16le') for n in keep) + b'\0\0',
        'VBA/dir': compress_raw(bytes(prefix) + b''.join(modules) + directory[pos:]),
        # MS-OVBA _VBA_PROJECT reserved header, no implementation-specific performance cache.
        'VBA/_VBA_PROJECT': bytes.fromhex('cc61ffff000000'),
    }
    expected_sources = {}
    for name, suffix in [('ThisWorkbook', 'cls'), ('Module1', 'bas')]:
        source = (ROOT / 'excel' / (name + '.' + suffix)).read_text().replace('\r\n', '\n').replace('\n', '\r\n')
        expected_sources[name] = source
        streams['VBA/' + name] = compress_raw(source.encode('cp1252'))
    target = ROOT / 'vendor/protocol/vbaProject.bin'
    target.write_bytes(write_cfb(streams))
    # Independently parse the generated compound document and decompress every module.
    parsed = {Path(name).stem: source for _, _, name, source in VBA_Parser(str(target)).extract_macros()}
    assert set(parsed) == set(expected_sources), parsed.keys()
    for name, source in expected_sources.items():
        assert parsed[name].rstrip() == source.rstrip(), name
    result = olefile.OleFileIO(target)
    assert not any('__SRP' in '/'.join(p) for p in result.listdir())
    print(target.relative_to(ROOT), len(target.read_bytes()), 'bytes; source modules independently verified')


if __name__ == '__main__':
    build()
