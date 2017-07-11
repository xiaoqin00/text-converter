import os
import mock
from itertools import chain
from lxml import etree
from copy import copy

from docx2html.core import (
    DEFAULT_LIST_NUMBERING_STYLE,
    _is_top_level_upper_roman,
    convert_image,
    create_html,
    get_font_size,
    get_image_id,
    get_single_list_nodes_data,
    get_ordered_list_type,
    get_namespace,
    get_relationship_info,
    get_style_dict,
    is_last_li,
)
from docx2html.tests.document_builder import DocxBuilder as DXB
from docx2html.tests import (
    _TranslationTestCase,
    assert_html_equal,
)


class SimpleListTestCase(_TranslationTestCase):
    expected_output = '''
        <html>
            <ol data-list-type="lower-alpha">
                <li>AAA</li>
                <li>BBB</li>
                <li>CCC</li>
            </ol>
        </html>
    '''

    # Ensure its not failing somewhere and falling back to decimal
    numbering_dict = {
        '1': {
            0: 'lowerLetter',
        }
    }

    def get_xml(self):
        li_text = [
            ('AAA', 0, 1),
            ('BBB', 0, 1),
            ('CCC', 0, 1),
        ]
        lis = ''
        for text, ilvl, numId in li_text:
            lis += DXB.li(text=text, ilvl=ilvl, numId=numId)

        xml = DXB.xml(lis)
        return etree.fromstring(xml)

    def test_get_li_nodes(self):
        tree = self.get_xml()
        meta_data = self.get_meta_data()
        w_namespace = get_namespace(tree, 'w')
        first_p_tag = tree.find('%sp' % w_namespace)

        li_data = get_single_list_nodes_data(first_p_tag, meta_data)
        assert len(list(li_data)) == 3

    def test_is_last_li(self):
        tree = self.get_xml()
        meta_data = self.get_meta_data()
        p_tags = tree.xpath('.//w:p', namespaces=tree.nsmap)
        result = [is_last_li(p, meta_data, current_numId='1') for p in p_tags]
        self.assertEqual(
            result,
            [False, False, True],
        )

    def test_get_list_type_valid(self):
        meta_data = self.get_meta_data()
        numId = '1'
        ilvl = 0

        # Show that for a valid combination of numId and ilvl that you get the
        # correct list type.
        list_type = get_ordered_list_type(meta_data, numId, ilvl)
        self.assertEqual(list_type, 'lowerLetter')

    def test_get_list_type_invalid_numId(self):
        meta_data = self.get_meta_data()
        numId = '2'  # Not valid
        ilvl = 0

        list_type = get_ordered_list_type(meta_data, numId, ilvl)
        self.assertEqual(list_type, DEFAULT_LIST_NUMBERING_STYLE)

    def test_get_list_type_invalid_ilvl(self):
        meta_data = self.get_meta_data()
        numId = '1'
        ilvl = 1  # Not valid

        list_type = get_ordered_list_type(meta_data, numId, ilvl)
        self.assertEqual(list_type, DEFAULT_LIST_NUMBERING_STYLE)

    def test_get_list_type_invalid_numId_and_ilvl(self):
        meta_data = self.get_meta_data()
        numId = '2'  # Not valid
        ilvl = 1  # Not valid

        list_type = get_ordered_list_type(meta_data, numId, ilvl)
        self.assertEqual(list_type, DEFAULT_LIST_NUMBERING_STYLE)


class TableInListTestCase(_TranslationTestCase):
    expected_output = '''
        <html>
            <ol data-list-type="decimal">
                <li>AAA<br />
                    <table>
                        <tr>
                            <td>BBB</td>
                            <td>CCC</td>
                        </tr>
                        <tr>
                            <td>DDD</td>
                            <td>EEE</td>
                        </tr>
                    </table>
                </li>
                <li>FFF</li>
            </ol>
            <p>GGG</p>
        </html>
    '''

    def get_xml(self):
        table = DXB.table(num_rows=2, num_columns=2, text=chain(
            [DXB.p_tag('BBB')],
            [DXB.p_tag('CCC')],
            [DXB.p_tag('DDD')],
            [DXB.p_tag('EEE')],
        ))

        # Nest that table in a list.
        first_li = DXB.li(text='AAA', ilvl=0, numId=1)
        second = DXB.li(text='FFF', ilvl=0, numId=1)
        p_tag = DXB.p_tag('GGG')

        body = ''
        for el in [first_li, table, second, p_tag]:
            body += el
        xml = DXB.xml(body)
        return etree.fromstring(xml)

    def test_get_li_nodes_with_nested_table(self):
        # Create a table
        tree = self.get_xml()
        meta_data = self.get_meta_data()
        w_namespace = get_namespace(tree, 'w')
        first_p_tag = tree.find('%sp' % w_namespace)

        # Show that list nesting deals with the table nesting
        li_data = get_single_list_nodes_data(first_p_tag, meta_data)
        assert len(list(li_data)) == 3

    def test_is_last_li(self):
        tree = self.get_xml()
        meta_data = self.get_meta_data()
        result = [is_last_li(el, meta_data, current_numId='1') for el in tree]
        self.assertEqual(
            result,
            # None list items are ignored
            [False, False, True, False],
        )


class RomanNumeralToHeadingTestCase(_TranslationTestCase):
    numbering_dict = {
        '1': {
            0: 'upperRoman',
            1: 'decimal',
            2: 'upperRoman',
        }
    }
    expected_output = '''
        <html>
            <h2>AAA</h2>
            <ol data-list-type="decimal">
                <li>BBB</li>
            </ol>
            <h2>CCC</h2>
            <ol data-list-type="decimal">
                <li>DDD</li>
            </ol>
            <h2>EEE</h2>
            <ol data-list-type="decimal">
                <li>FFF
                    <ol data-list-type="upper-roman">
                        <li>GGG</li>
                    </ol>
                </li>
            </ol>
        </html>
    '''

    def get_xml(self):
        li_text = [
            ('AAA', 0, 1),
            ('BBB', 1, 1),
            ('CCC', 0, 1),
            ('DDD', 1, 1),
            ('EEE', 0, 1),
            ('FFF', 1, 1),
            ('GGG', 2, 1),
        ]
        body = ''
        for text, ilvl, numId in li_text:
            body += DXB.li(text=text, ilvl=ilvl, numId=numId)

        xml = DXB.xml(body)
        return etree.fromstring(xml)

    def test_is_top_level_upper_roman(self):
        tree = self.get_xml()
        w_namespace = get_namespace(tree, 'w')
        meta_data = self.get_meta_data()

        result = []
        for p in tree.findall('%sp' % w_namespace):
            result.append(
                _is_top_level_upper_roman(p, meta_data)
            )
        self.assertEqual(
            result,
            [
                True,  # AAA
                False,  # BBB
                True,  # CCC
                False,  # DDD
                True,  # EEE
                False,  # FFF
                False,  # GGG - notice this is upper roman but not in the root
            ]
        )


class RomanNumeralToHeadingAllBoldTestCase(_TranslationTestCase):
    numbering_dict = {
        '1': {
            0: 'upperRoman',
        }
    }
    expected_output = '''
        <html>
            <h2>AAA</h2>
            <h2>BBB</h2>
            <h2>CCC</h2>
        </html>
    '''

    def get_xml(self):
        li_text = [
            ('AAA', 0, 1),
            ('BBB', 0, 1),
            ('CCC', 0, 1),
        ]
        body = ''
        for text, ilvl, numId in li_text:
            body += DXB.li(text=text, ilvl=ilvl, numId=numId, bold=True)

        xml = DXB.xml(body)
        return etree.fromstring(xml)


class ImageTestCase(_TranslationTestCase):
    relationship_dict = {
        'rId0': 'media/image1.jpeg',
        'rId1': 'media/image2.jpeg',
    }
    image_sizes = {
        'rId0': (4, 4),
        'rId1': (4, 4),
    }
    expected_output = '''
        <html>
            <p>
                <img src="media/image1.jpeg" height="4" width="4" />
            </p>
            <p>
                <img src="media/image2.jpeg" height="4" width="4" />
            </p>
        </html>
    '''

    @staticmethod
    def image_handler(image_id, relationship_dict):
        return relationship_dict.get(image_id)

    def get_xml(self):
        drawing = DXB.drawing('rId0')
        pict = DXB.pict('rId1')
        tags = [
            drawing,
            pict,
        ]
        body = ''
        for el in tags:
            body += el

        xml = DXB.xml(body)
        return etree.fromstring(xml)

    def test_get_image_id(self):
        tree = self.get_xml()
        els = []
        w_namespace = get_namespace(tree, 'w')
        for el in tree.iter():
            if el.tag == '%sdrawing' % w_namespace:
                els.append(el)
            if el.tag == '%spict' % w_namespace:
                els.append(el)
        image_ids = []
        for el in els:
            image_ids.append(get_image_id(el))
        self.assertEqual(
            image_ids,
            [
                'rId0',
                'rId1',
            ]
        )

    @mock.patch('docx2html.core._get_image_size_from_image')
    def test_missing_size(self, patched_item):
        def side_effect(*args, **kwargs):
            return (6, 6)
        patched_item.side_effect = side_effect
        tree = self.get_xml()
        meta_data = copy(self.get_meta_data())
        del meta_data.image_sizes['rId1']

        html = create_html(tree, meta_data)

        # Show that the height and width were grabbed from the actual image.
        assert_html_equal(html, '''
            <html>
                <p>
                    <img src="media/image1.jpeg" height="4" width="4" />
                </p>
                <p>
                    <img src="media/image2.jpeg" height="6" width="6" />
                </p>
            </html>
        ''')


class SkipImageTestCase(_TranslationTestCase):
    relationship_dict = {
        # These are only commented out because ``get_relationship_info`` strips
        # them out, however since we have image_sizes I want to show that they
        # are intentionally not added to the ``relationship_dict``
        #'rId0': 'media/image1.svg',
        #'rId1': 'media/image2.emf',
        #'rId2': 'media/image3.wmf',
    }
    image_sizes = {
        'rId0': (4, 4),
        'rId1': (4, 4),
        'rId2': (4, 4),
    }
    expected_output = '<html></html>'

    @staticmethod
    def image_handler(image_id, relationship_dict):
        return relationship_dict.get(image_id)

    def get_xml(self):
        tags = [
            DXB.drawing('rId2'),
            DXB.drawing('rId3'),
            DXB.drawing('rId4'),
        ]
        body = ''
        for el in tags:
            body += el

        xml = DXB.xml(body)
        return etree.fromstring(xml)

    def test_get_relationship_info(self):
        tree = self.get_xml()
        media = {
            'media/image1.svg': 'test',
            'media/image2.emf': 'test',
            'media/image3.wmf': 'test',
        }
        relationship_info = get_relationship_info(
            tree,
            media,
            self.image_sizes,
        )
        self.assertEqual(relationship_info, {})


class ImageNoSizeTestCase(_TranslationTestCase):
    relationship_dict = {
        'rId0': os.path.join(
            os.path.abspath(os.path.dirname(__file__)),
            '..',
            'fixtures',
            'bullet_go_gray.png',
        )
    }
    image_sizes = {
        'rId0': (0, 0),
    }
    expected_output = '''
        <html>
            <p>
                <img src="%s" />
            </p>
        </html>
    ''' % relationship_dict['rId0']

    @staticmethod
    def image_handler(image_id, relationship_dict):
        return relationship_dict.get(image_id)

    def get_xml(self):
        drawing = DXB.drawing('rId0')
        tags = [
            drawing,
        ]
        body = ''
        for el in tags:
            body += el

        xml = DXB.xml(body)
        return etree.fromstring(xml)

    def test_convert_image(self):
        convert_image(self.relationship_dict['rId0'], self.image_sizes['rId0'])


class ListWithContinuationTestCase(_TranslationTestCase):
    expected_output = '''
        <html>
            <ol data-list-type="decimal">
                <li>AAA<br />BBB</li>
                <li>CCC<br />
                    <table>
                        <tr>
                            <td>DDD</td>
                            <td>EEE</td>
                        </tr>
                        <tr>
                            <td>FFF</td>
                            <td>GGG</td>
                        </tr>
                    </table>
                </li>
                <li>HHH</li>
            </ol>
        </html>
    '''

    def get_xml(self):
        table = DXB.table(num_rows=2, num_columns=2, text=chain(
            [DXB.p_tag('DDD')],
            [DXB.p_tag('EEE')],
            [DXB.p_tag('FFF')],
            [DXB.p_tag('GGG')],
        ))
        tags = [
            DXB.li(text='AAA', ilvl=0, numId=1),
            DXB.p_tag('BBB'),
            DXB.li(text='CCC', ilvl=0, numId=1),
            table,
            DXB.li(text='HHH', ilvl=0, numId=1),
        ]
        body = ''
        for el in tags:
            body += el

        xml = DXB.xml(body)
        return etree.fromstring(xml)


class PictImageTestCase(_TranslationTestCase):
    relationship_dict = {
        'rId0': 'media/image1.jpeg',
    }
    image_sizes = {
        'rId0': (4, 4),
    }
    expected_output = '''
        <html>
            <p>
                <img src="media/image1.jpeg" height="4" width="4" />
            </p>
        </html>
    '''

    @staticmethod
    def image_handler(image_id, relationship_dict):
        return relationship_dict.get(image_id)

    def get_xml(self):
        pict = DXB.pict('rId0')
        tags = [
            pict,
        ]
        body = ''
        for el in tags:
            body += el

        xml = DXB.xml(body)
        return etree.fromstring(xml)

    def test_image_id_for_pict(self):
        tree = self.get_xml()

        # Get all the pict tags
        pict_tags = tree.xpath('.//w:pict', namespaces=tree.nsmap)
        self.assertEqual(len(pict_tags), 1)

        # Get the image id for the pict tag.
        pict_tag = pict_tags[0]
        image_id = get_image_id(pict_tag)
        self.assertEqual(image_id, 'rId0')


class PictImageMissingIdTestCase(_TranslationTestCase):
    expected_output = '''
        <html></html>
    '''

    def get_xml(self):
        pict = DXB.pict(None)
        tags = [
            pict,
        ]
        body = ''
        for el in tags:
            body += el

        xml = DXB.xml(body)
        return etree.fromstring(xml)


class TableWithInvalidTag(_TranslationTestCase):
    expected_output = '''
        <html>
            <table>
                <tr>
                    <td>AAA</td>
                    <td>BBB</td>
                </tr>
                <tr>
                    <td></td>
                    <td>DDD</td>
                </tr>
            </table>
        </html>
    '''

    def get_xml(self):
        table = DXB.table(num_rows=2, num_columns=2, text=chain(
            [DXB.p_tag('AAA')],
            [DXB.p_tag('BBB')],
            # This tag may have CCC in it, however this tag has no meaning
            # pertaining to content.
            ['<w:invalidTag>CCC</w:invalidTag>'],
            [DXB.p_tag('DDD')],
        ))
        body = table
        xml = DXB.xml(body)
        return etree.fromstring(xml)


class NonStandardTextTagsTestCase(_TranslationTestCase):
    expected_output = '''
    <html>
        <p>insert smarttag</p>
    </html>
    '''

    def get_xml(self):
        run_tags = [DXB.r_tag(i) for i in 'insert ']
        insert_tag = DXB.insert_tag(run_tags)
        run_tags = [DXB.r_tag(i) for i in 'smarttag']
        smart_tag = DXB.smart_tag(run_tags)

        run_tags = [insert_tag, smart_tag]
        body = DXB.p_tag(run_tags)
        xml = DXB.xml(body)
        return etree.fromstring(xml)


class HyperlinkStyledTestCase(_TranslationTestCase):
    relationship_dict = {
        'rId0': 'www.google.com',
    }

    expected_output = '''
    <html>
        <p><a href="www.google.com">link</a>.</p>
    </html>
    '''

    def get_xml(self):
        run_tags = []
        run_tags.append(DXB.r_tag('link', is_bold=True))
        run_tags = [DXB.hyperlink_tag(r_id='rId0', run_tags=run_tags)]
        run_tags.append(DXB.r_tag('.', is_bold=False))
        body = DXB.p_tag(run_tags)
        xml = DXB.xml(body)
        return etree.fromstring(xml)


class HyperlinkWithMultipleRunsTestCase(_TranslationTestCase):
    relationship_dict = {
        'rId0': 'www.google.com',
    }

    expected_output = '''
    <html>
        <p><a href="www.google.com">link</a>.</p>
    </html>
    '''

    def get_xml(self):
        run_tags = [DXB.r_tag(i) for i in 'link']
        run_tags = [DXB.hyperlink_tag(r_id='rId0', run_tags=run_tags)]
        run_tags.append(DXB.r_tag('.', is_bold=False))
        body = DXB.p_tag(run_tags)
        xml = DXB.xml(body)
        return etree.fromstring(xml)


class HyperlinkNoTextTestCase(_TranslationTestCase):
    relationship_dict = {
        'rId0': 'www.google.com',
    }

    expected_output = '''
    <html>
    </html>
    '''

    def get_xml(self):
        run_tags = []
        run_tags = [DXB.hyperlink_tag(r_id='rId0', run_tags=run_tags)]
        body = DXB.p_tag(run_tags)
        xml = DXB.xml(body)
        return etree.fromstring(xml)


class HyperlinkVanillaTestCase(_TranslationTestCase):
    relationship_dict = {
        'rId0': 'www.google.com',
    }

    expected_output = '''
    <html>
        <p><a href="www.google.com">link</a>.</p>
    </html>
    '''

    def get_xml(self):
        run_tags = []
        run_tags.append(DXB.r_tag('link', is_bold=False))
        run_tags = [DXB.hyperlink_tag(r_id='rId0', run_tags=run_tags)]
        run_tags.append(DXB.r_tag('.', is_bold=False))
        body = DXB.p_tag(run_tags)
        xml = DXB.xml(body)
        return etree.fromstring(xml)


class HyperlinkWithBreakTestCase(_TranslationTestCase):
    relationship_dict = {
        'rId0': 'www.google.com',
    }

    expected_output = '''
    <html>
        <p><a href="www.google.com">link<br /></a></p>
    </html>
    '''

    def get_xml(self):
        run_tags = []
        run_tags.append(DXB.r_tag('link'))
        run_tags.append(DXB.r_tag(None, include_linebreak=True))
        run_tags = [DXB.hyperlink_tag(r_id='rId0', run_tags=run_tags)]
        body = DXB.p_tag(run_tags)
        xml = DXB.xml(body)
        return etree.fromstring(xml)


class MissingFontInfoTestCase(_TranslationTestCase):
    styles_dict = {
        'BodyText': {
            'header': False, 'font_size': None, 'based_on': 'Normal',
        },
    }

    expected_output = '''
    <html>
        <p><strong>AAA</strong></p>
    </html>
    '''

    def get_xml(self):
        p_tag = '''
        <w:p w:rsidR="009C063D" w:rsidRDefault="009C063D">
            <w:pPr>
                <w:pStyle w:val="BodyText"/>
                <w:ind w:left="720"/>
                <w:rPr>
                    <w:b w:val="0"/>
                    <w:bCs w:val="0"/>
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:b w:val="0"/>
                    <w:bCs w:val="0"/>
                </w:rPr>
                <w:t>AAA</w:t>
            </w:r>
        </w:p>
        '''
        xml = DXB.xml(p_tag)
        return etree.fromstring(xml)

    def test_get_font_size(self):
        tree = self.get_xml()
        w_namespace = get_namespace(tree, 'w')
        p_tag = tree.find('%sp' % w_namespace)
        self.assertNotEqual(p_tag, None)
        self.assertEqual(
            get_font_size(p_tag, self.styles_dict),
            None,
        )

    def test_get_font_size_empty_styles_dict(self):
        tree = self.get_xml()
        w_namespace = get_namespace(tree, 'w')
        p_tag = tree.find('%sp' % w_namespace)
        self.assertNotEqual(p_tag, None)
        self.assertEqual(
            get_font_size(p_tag, {}),
            None,
        )


class HeaderFooterTagsWithContent(_TranslationTestCase):
    expected_output = '''
    <html>
        <ol data-list-type="decimal">
            <li>AAA</li>
        </ol>
    </html>
    '''

    def get_xml(self):
        li = DXB.li(text='AAA', ilvl=0, numId=1)
        p_tag = DXB.p_tag('BBB')
        footer_tag = DXB.sectPr_tag(p_tag)
        body = li + footer_tag
        xml = DXB.xml(body)
        return etree.fromstring(xml)


class StylesParsingTestCase(_TranslationTestCase):
    expected_output = '<html></html>'

    def get_xml(self):
        return etree.fromstring(DXB.xml(''))

    def test_get_headings(self):

        styles = [
            DXB.style('heading 1', 'heading 1'),
        ]
        xml = DXB.styles_xml(styles)
        styles_xml = etree.fromstring(xml)
        styles_dict = get_style_dict(styles_xml)
        self.assertEqual(styles_dict['heading 1']['header'], 'h2')


class MangledIlvlTestCase(_TranslationTestCase):
    expected_output = '''
    <html>
        <ol data-list-type="decimal">
            <li>AAA</li>
        </ol>
        <ol data-list-type="decimal">
            <li>BBB</li>
        </ol>
        <ol data-list-type="decimal">
            <li>CCC</li>
        </ol>
    </html>
    '''

    def get_xml(self):
        li_text = [
            ('AAA', 0, 2),
            ('BBB', 1, 1),
            ('CCC', 0, 1),
        ]
        lis = ''
        for text, ilvl, numId in li_text:
            lis += DXB.li(text=text, ilvl=ilvl, numId=numId)

        xml = DXB.xml(lis)
        return etree.fromstring(xml)


class SeperateListsTestCase(_TranslationTestCase):
    expected_output = '''
    <html>
        <ol data-list-type="decimal">
            <li>AAA</li>
        </ol>
        <ol data-list-type="decimal">
            <li>BBB</li>
        </ol>
        <ol data-list-type="decimal">
            <li>CCC</li>
        </ol>
    </html>
    '''

    def get_xml(self):
        li_text = [
            ('AAA', 0, 2),
            # Because AAA and CCC are part of the same list (same list id)
            # and BBB is different, these need to be split into three
            # lists (or lose everything from BBB and after.
            ('BBB', 0, 1),
            ('CCC', 0, 2),
        ]
        lis = ''
        for text, ilvl, numId in li_text:
            lis += DXB.li(text=text, ilvl=ilvl, numId=numId)

        xml = DXB.xml(lis)
        return etree.fromstring(xml)


class InvalidIlvlOrderTestCase(_TranslationTestCase):
    expected_output = '''
    <html>
        <ol data-list-type="decimal">
            <li>AAA
                <ol data-list-type="decimal">
                    <li>BBB</li>
                </ol>
                <ol data-list-type="decimal">
                    <li>CCC</li>
                </ol>
            </li>
        </ol>
    </html>
    '''

    def get_xml(self):
        tags = [
            DXB.li(text='AAA', ilvl=1, numId=1),
            DXB.li(text='BBB', ilvl=3, numId=1),
            DXB.li(text='CCC', ilvl=2, numId=1),
        ]
        body = ''
        for el in tags:
            body += el

        xml = DXB.xml(body)
        return etree.fromstring(xml)


class RTagWithNoText(_TranslationTestCase):
    expected_output = '<html></html>'

    def get_xml(self):
        p_tag = DXB.p_tag(None)  # No text
        run_tags = [p_tag]
        # The bug is only present in a hyperlink
        run_tags = [DXB.hyperlink_tag(r_id='rId0', run_tags=run_tags)]
        body = DXB.p_tag(run_tags)

        xml = DXB.xml(body)
        return etree.fromstring(xml)
