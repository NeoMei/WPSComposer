import importlib.util
from pathlib import Path
import unittest
import xml.etree.ElementTree as ET
ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('configure',ROOT/'configure.py')
configure=importlib.util.module_from_spec(spec);spec.loader.exec_module(configure)
NS={'m':'http://schemas.microsoft.com/office/appforoffice/1.1'}
class ManifestTests(unittest.TestCase):
    def test_word_only_minimum_permission_and_https_origin(self):
        value=configure.render('https://localhost:3443')
        root=ET.fromstring(value)
        self.assertEqual(root.attrib['{http://www.w3.org/2001/XMLSchema-instance}type'],'TaskPaneApp')
        self.assertEqual([n.attrib['Name'] for n in root.findall('m:Hosts/m:Host',NS)],['Document'])
        self.assertEqual(root.find('m:Permissions',NS).text,'ReadWriteDocument')
        self.assertEqual(root.find('m:DefaultSettings/m:SourceLocation',NS).attrib['DefaultValue'],'https://localhost:3443/taskpane.html')
        self.assertEqual([(n.attrib['Name'],n.attrib['MinVersion']) for n in root.findall('m:Requirements/m:Sets/m:Set',NS)],[('WordApi','1.1')])
        self.assertNotIn('VersionOverrides',value)
    def test_reject_non_loopback_or_unsafe_origins(self):
        for value in ['http://localhost:3443','https://example.com','https://localhost.evil.test','https://user:pass@localhost','https://localhost/path','https://localhost/?q=x','https://localhost/#x','https://127.0.0.2:3443','https://localhost:0','https://localhost:65536','https://localhost:443"&x']:
            with self.subTest(value=value),self.assertRaises(ValueError):configure.render(value)
    def test_explicit_ipv4_and_ipv6_loopback(self):
        for value in ['https://127.0.0.1:3443','https://[::1]:3443']:
            self.assertIn(value+'/taskpane.html',configure.render(value))
    def test_only_fixed_output_in_package_and_no_clobber(self):
        self.assertEqual(configure.OUTPUT.parent,ROOT)
        self.assertEqual(configure.OUTPUT.name,'manifest.word.xml')
    def test_source_has_only_official_remote_script(self):
        from html.parser import HTMLParser
        class Scripts(HTMLParser):
            def __init__(self):super().__init__();self.src=[]
            def handle_starttag(self,tag,attrs):
                if tag=='script':self.src.append(dict(attrs).get('src'))
        parser=Scripts();parser.feed((ROOT/'taskpane.html').read_text())
        remote=[s for s in parser.src if s and s.startswith('https:')]
        self.assertEqual(remote,['https://appsforoffice.microsoft.com/lib/1/hosted/office.js'])
        self.assertTrue(all(parser.src))
    def test_no_transport_persistence_or_document_mutation(self):
        import re
        source='\n'.join((ROOT/p).read_text() for p in ('probe.js','taskpane.js'))
        for forbidden in [r'\bfetch\s*\(',r'new\s+WebSocket',r'XMLHttpRequest',r'sendBeacon',r'localStorage',r'sessionStorage',r'\.settings\b',r'customXmlParts',r'\.insert\w+\s*\(',r'\.getFileAsync\s*\(',r'\.setSelectedDataAsync\s*\(',r'\.save\s*\(',r'\beval\s*\(']:
            self.assertIsNone(re.search(forbidden,source),forbidden)
if __name__=='__main__':unittest.main()
