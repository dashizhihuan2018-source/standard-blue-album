import re
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]


class PublicAssetsTest(unittest.TestCase):
    def test_sources_do_not_contain_user_home_or_credentials(self):
        patterns = [r'/Users/[A-Za-z0-9_.-]+/', r'/home/[A-Za-z0-9_.-]+/',
                    r'-----BEGIN [A-Z ]*PRIVATE KEY-----',
                    r'gh[pousr]_[A-Za-z0-9]{30,}', r'sk-[A-Za-z0-9]{30,}']
        for folder in ('references', 'samples', 'tests'):
            for path in (ROOT/folder).glob('*.py'):
                with self.subTest(file=path.name):
                    text = path.read_text()
                    for pattern in patterns:
                        self.assertIsNone(re.search(pattern, text))

    def test_pptx_examples_have_neutral_metadata_and_no_comments(self):
        for path in (ROOT/'samples').glob('*.pptx'):
            with self.subTest(file=path.name), zipfile.ZipFile(path) as package:
                props = ET.fromstring(package.read('docProps/core.xml'))
                author = props.find('{http://purl.org/dc/elements/1.1/}creator')
                self.assertIn(author.text or '', ('示例机构', ''))
                editor = props.find('{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy')
                self.assertFalse(editor is not None and editor.text)
                self.assertFalse(any(n.startswith(('ppt/comments/', 'ppt/notesSlides/', 'ppt/embeddings/')) for n in package.namelist()))


if __name__ == '__main__':
    unittest.main()
