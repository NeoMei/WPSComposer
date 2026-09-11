"""Adversarial native-artifact checks using only in-memory XML/package doubles."""
import xml.etree.ElementTree as ET
import pytest
from fixtures.microsoft_parity import macos_word_business_followup as fixture

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P='http://schemas.openxmlformats.org/package/2006/relationships'


def simple(command='PAGE'):
    return f'<w:fldSimple w:instr="{command}"><w:r><w:t>1</w:t></w:r></w:fldSimple>'


def complex_field(command='PAGE'):
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText> {command} </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')


class MemoryPackage:
    def __init__(self, footer, *, body_field='', references=('r1', None), orphan=None, target_mode=None):
        sections=''.join('<w:p><w:pPr><w:sectPr>'+(
            f'<w:footerReference w:type="default" r:id="{ref}"/>' if ref else '')+
            '</w:sectPr></w:pPr></w:p>' for ref in references)
        self.body=ET.fromstring(f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body_field}{sections}</w:body></w:document>')
        mode=f' TargetMode="{target_mode}"' if target_mode else ''
        self.parts={'word/footer1.xml':f'<w:ftr xmlns:w="{W}"><w:p>{footer}</w:p></w:ftr>'.encode(),
            'word/_rels/document.xml.rels':f'<Relationships xmlns="{P}"><Relationship Id="r1" Type="{R}/footer" Target="footer1.xml"{mode}/></Relationships>'.encode()}
        if orphan is not None:self.parts['word/footer99.xml']=f'<w:ftr xmlns:w="{W}"><w:p>{orphan}</w:p></w:ftr>'.encode()
    def read(self,name):return self.parts[name]
    def namelist(self):return list(self.parts)


@pytest.mark.parametrize('field',[simple(),complex_field(),complex_field('PAGE \\* MERGEFORMAT')])
def test_real_page_field_in_effective_inherited_default_footer_passes(field):
    p=MemoryPackage(field)
    fixture._verify_page_fields(p,p.body)


@pytest.mark.parametrize('field',[
    '<w:r><w:t>PAGE</w:t></w:r>',simple('NUMPAGES'),simple('PAGEREF Bookmark'),
    complex_field('NUMPAGES'),complex_field('PAGEREF Bookmark'),
    complex_field().replace('<w:fldChar w:fldCharType="end"/>',''),
    complex_field().replace('<w:fldChar w:fldCharType="begin"/>',''),
    complex_field().replace('<w:fldChar w:fldCharType="separate"/>',''),
    '<w:r><w:instrText>PAGE</w:instrText></w:r>',
])
def test_footer_plaintext_other_commands_and_unbalanced_fields_rejected(field):
    p=MemoryPackage(field)
    with pytest.raises(AssertionError):fixture._verify_page_fields(p,p.body)


@pytest.mark.parametrize('body_field',[simple(),complex_field()])
def test_body_page_fields_rejected_in_both_representations(body_field):
    p=MemoryPackage(complex_field(),body_field=body_field)
    with pytest.raises(AssertionError,match='body'):fixture._verify_page_fields(p,p.body)


def test_unreferenced_page_footer_does_not_certify_referenced_plaintext_footer():
    p=MemoryPackage('<w:r><w:t>Page 1</w:t></w:r>',orphan=complex_field())
    with pytest.raises(AssertionError):fixture._verify_page_fields(p,p.body)


@pytest.mark.parametrize('kwargs',[{'references':(None,None)},{'references':('unknown',None)},{'target_mode':'External'}])
def test_missing_or_external_default_footer_binding_rejected(kwargs):
    p=MemoryPackage(complex_field(),**kwargs)
    with pytest.raises(AssertionError):fixture._verify_page_fields(p,p.body)


def test_instruction_split_across_runs_is_reassembled_without_matching_result_text():
    field=complex_field().replace(' PAGE ', ' PA</w:instrText></w:r><w:r><w:instrText>GE ')
    p=MemoryPackage(field)
    assert fixture._verify_page_fields(p,p.body)==['word/footer1.xml','word/footer1.xml']
    p=MemoryPackage(complex_field('NUMPAGES').replace('<w:t>1</w:t>','<w:t>PAGE</w:t>'))
    with pytest.raises(AssertionError):fixture._verify_page_fields(p,p.body)


def test_even_footer_page_field_cannot_replace_missing_default_footer():
    p=MemoryPackage(complex_field())
    for node in p.body.iter('{'+W+'}footerReference'):node.set('{'+W+'}type','even')
    with pytest.raises(AssertionError):fixture._verify_page_fields(p,p.body)


def test_later_section_explicit_default_footer_overrides_previous_page_field():
    p=MemoryPackage(complex_field(),references=('r1','r2'))
    rels=ET.fromstring(p.parts['word/_rels/document.xml.rels'])
    ET.SubElement(rels,'{'+P+'}Relationship',{'Id':'r2','Type':R+'/footer','Target':'footer2.xml'})
    p.parts['word/_rels/document.xml.rels']=ET.tostring(rels)
    p.parts['word/footer2.xml']=f'<w:ftr xmlns:w="{W}"><w:p><w:r><w:t>PAGE</w:t></w:r></w:p></w:ftr>'.encode()
    with pytest.raises(AssertionError):fixture._verify_page_fields(p,p.body)
