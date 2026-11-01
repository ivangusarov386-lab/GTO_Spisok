#!/usr/bin/env python3
"""Regression checks for CFB/OVBA structure; not a Microsoft Excel execution test."""
import unittest
from pathlib import Path
import olefile
from oletools.olevba import VBA_Parser, decompress_stream
from build_vba import compress_vba

ROOT = Path(__file__).resolve().parent.parent
PROJECT = ROOT / 'vendor/protocol/vbaProject.bin'


class VbaStructureTests(unittest.TestCase):
    def test_compression_never_appends_padding(self):
        for size in [0, 1, 7, 8, 9, 1024, 3640, 4095, 4096, 4097, 8192]:
            source = (b'Attribute VB_Name = "Module1"\r\nOption Explicit\r\n' * 200)[:size]
            with self.subTest(size=size):
                self.assertEqual(bytes(decompress_stream(compress_vba(source))), source)

    def test_storage_start_sector_is_zero(self):
        with olefile.OleFileIO(PROJECT, raise_defects=olefile.DEFECT_INCORRECT) as project:
            storages = [e for e in project.direntries if e and e.entry_type == 1]
            self.assertTrue(storages)
            for entry in storages:
                self.assertEqual(entry.isectStart, 0, entry.name)
                self.assertEqual(entry.size, 0, entry.name)

    def test_directory_ends_at_terminator(self):
        with olefile.OleFileIO(PROJECT) as project:
            directory = decompress_stream(project.openstream('VBA/dir').read())
            self.assertTrue(directory.endswith(b'\x10\x00\x00\x00\x00\x00'))

    def test_sources_are_exact_not_space_padded(self):
        parser = VBA_Parser(str(PROJECT))
        extracted = {name: code for _, _, name, code in parser.extract_macros()}
        parser.close()
        self.assertEqual(set(extracted), {'ThisWorkbook.cls', 'Module1.bas'})
        for name, code in extracted.items():
            expected = (ROOT / 'excel' / name).read_text().replace('\r\n', '\n').replace('\n', '\r\n')
            self.assertEqual(code, expected, name)


if __name__ == '__main__':
    unittest.main()
