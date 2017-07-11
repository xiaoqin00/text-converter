import cgi
import logging
import os
import os.path
import re
from PIL import Image
from lxml import etree
from lxml.etree import XMLSyntaxError

from collections import namedtuple, defaultdict
from zipfile import ZipFile, BadZipfile

from docx2html.exceptions import (
    ConversionFailed,
    FileNotDocx,
    MalformedDocx,
    UnintendedTag,
    SyntaxNotSupported,
)

DETECT_FONT_SIZE = False
EMUS_PER_PIXEL = 9525
NSMAP = {}
IMAGE_EXTENSIONS_TO_SKIP = ['emf', 'wmf', 'svg']
DEFAULT_LIST_NUMBERING_STYLE = 'decimal'

logger = logging.getLogger(__name__)

###
# Help functions
###


def replace_ext(file_path, new_ext):
    """
    >>> replace_ext('one/two/three.four.doc', '.html')
    'one/two/three.four.html'
    >>> replace_ext('one/two/three.four.DOC', '.html')
    'one/two/three.four.html'
    >>> replace_ext('one/two/three.four.DOC', 'html')
    'one/two/three.four.html'
    """
    if not new_ext.startswith(os.extsep):
        new_ext = os.extsep + new_ext
    index = file_path.rfind(os.extsep)
    return file_path[:index] + new_ext


def ensure_tag(tags):
    # For some functions we can short-circuit and early exit if the tag is not
    # the right kind.

    def wrapped(f):
        def wrap(*args, **kwargs):
            passed_in_tag = args[0]
            if passed_in_tag is None:
                return None
            w_namespace = get_namespace(passed_in_tag, 'w')
            valid_tags = [
                '%s%s' % (w_namespace, t) for t in tags
            ]
            if passed_in_tag.tag in valid_tags:
                return f(*args, **kwargs)
            return None
        return wrap
    return wrapped


def get_namespace(el, namespace):
    if namespace not in NSMAP:
        NSMAP[namespace] = '{%s}' % el.nsmap[namespace]
    return NSMAP[namespace]


def convert_image(target, image_size):
    _, extension = os.path.splitext(os.path.basename(target))
    # If the image size has a zero in it early return
    if image_size and not all(image_size):
        return target
    # All the image types need to be converted to gif.
    invalid_extensions = (
        '.bmp',
        '.dib',
        '.tiff',
        '.tif',
    )
    # Open the image and get the format.
    try:
        image = Image.open(target)
    except IOError:
        return target
    image_format = image.format
    image_file_name = target

    # Make sure the size of the image and the size of the embedded image are
    # the same.
    if image_size is not None and image.size != image_size:
        # Resize if needed
        try:
            image = image.resize(image_size, Image.ANTIALIAS)
        except IOError:
            pass

    # If we have an invalid extension, change the format to gif.
    if extension.lower() in invalid_extensions:
        image_format = 'GIF'
        image_file_name = replace_ext(target, '.gif')

    # Resave the image (Post resizing) with the correct format
    try:
        image.save(image_file_name, image_format)
    except IOError:
        return target
    return image_file_name


@ensure_tag(['p'])
def get_font_size(p, styles_dict):
    w_namespace = get_namespace(p, 'w')
    r = p.find('%sr' % w_namespace)
    if r is None:
        return None
    rpr = r.find('%srPr' % w_namespace)
    if rpr is None:
        return None
    size = rpr.find('%ssz' % w_namespace)
    if size is None:
        # Need to get the font size off the styleId
        pPr = p.find('%spPr' % w_namespace)
        if pPr is None:
            return None
        pStyle = pPr.find('%spStyle' % w_namespace)
        if pStyle is None:
            return None
        pStyle = pStyle.get('%sval' % w_namespace)
        font_size = None
        style_value = styles_dict.get(pStyle, None)
        if style_value is None:
            return None
        if 'font_size' in style_value:
            font_size = styles_dict[pStyle]['font_size']
        while font_size is None:
            old_pStyle = pStyle
            # If pStyle is not in the styles_dict then we have to break.
            if pStyle not in styles_dict:
                break
            # If based on is not in the styles_dict for pStyle then we have to
            # break.
            if 'based_on' not in styles_dict[pStyle]:
                break
            # Try to derive what the font size is based on what the current
            # style is based on.
            pStyle = styles_dict[pStyle]['based_on']
            if old_pStyle == pStyle:
                break
            # If pStyle is not in styles_dict then break.
            if pStyle not in styles_dict:
                break
            # We have found a new font size
            font_size = styles_dict[pStyle]['font_size']
        return font_size

    return size.get('%sval' % w_namespace)


@ensure_tag(['p'])
def is_natural_header(el, styles_dict):
    w_namespace = get_namespace(el, 'w')
    pPr = el.find('%spPr' % w_namespace)
    if pPr is None:
        return False
    pStyle = pPr.find('%spStyle' % w_namespace)
    if pStyle is None:
        return False
    style_id = pStyle.get('%sval' % w_namespace)
    if (
            style_id in styles_dict and
            'header' in styles_dict[style_id] and
            styles_dict[style_id]['header']):
        return styles_dict[style_id]['header']


@ensure_tag(['p'])
def is_header(el, meta_data):
    if _is_top_level_upper_roman(el, meta_data):
        return 'h2'
    el_is_natural_header = is_natural_header(el, meta_data.styles_dict)
    if el_is_natural_header:
        return el_is_natural_header
    if _is_li(el):
        return False
    w_namespace = get_namespace(el, 'w')
    if el.tag == '%stbl' % w_namespace:
        return False

    # Check to see if this is a header because the font size is different than
    # the normal font size.
    # Since get_font_size is a method used before meta is created, just pass in
    # styles_dict.
    if DETECT_FONT_SIZE:
        font_size = get_font_size(el, meta_data.styles_dict)
        if font_size is not None:
            if meta_data.font_sizes_dict[font_size]:
                return meta_data.font_sizes_dict[font_size]

    # If a paragraph is longer than eight words it is likely not supposed to be
    # an h tag.
    num_words = len(
        etree.tostring(
            el,
            encoding=unicode,
            method='text',
        ).split(' ')
    )
    if num_words > 8:
        return False

    # Check to see if the full line is bold.
    whole_line_bold, whole_line_italics = whole_line_styled(el)
    if whole_line_bold or whole_line_italics:
        return 'h2'

    return False


@ensure_tag(['p'])
def _is_top_level_upper_roman(el, meta_data):
    w_namespace = get_namespace(el, 'w')
    ilvl = get_ilvl(el, w_namespace)
    # If this list is not in the root document (indentation of 0), then it
    # cannot be a top level upper roman list.
    if ilvl != 0:
        return False
    numId = get_numId(el, w_namespace)
    list_type = meta_data.numbering_dict[numId].get(ilvl, False)
    return list_type == 'upperRoman'


@ensure_tag(['p'])
def _is_li(el):
    return len(el.xpath('.//w:numPr/w:ilvl', namespaces=el.nsmap)) != 0


@ensure_tag(['p'])
def is_li(el, meta_data):
    """
    The only real distinction between an ``li`` tag and a ``p`` tag is that an
    ``li`` tag has an attribute called numPr which holds the list id and ilvl
    (indentation level)
    """

    if is_header(el, meta_data):
        return False
    return _is_li(el)


def has_text(p):
    """
    It is possible for a ``p`` tag in document.xml to not have any content. If
    this is the case we do not want that tag interfering with things like
    lists. Detect if this tag has any content.
    """
    return '' != etree.tostring(p, encoding=unicode, method='text').strip()


def is_last_li(li, meta_data, current_numId):
    """
    Determine if ``li`` is the last list item for a given list
    """
    if not is_li(li, meta_data):
        return False
    w_namespace = get_namespace(li, 'w')
    next_el = li
    while True:
        # If we run out of element this must be the last list item
        if next_el is None:
            return True

        next_el = next_el.getnext()
        # Ignore elements that are not a list item
        if not is_li(next_el, meta_data):
            continue

        new_numId = get_numId(next_el, w_namespace)
        if current_numId != new_numId:
            return True
        # If we have gotten here then we have found another list item in the
        # current list, so ``li`` is not the last li in the list.
        return False


@ensure_tag(['p'])
def get_single_list_nodes_data(li, meta_data):
    """
    Find consecutive li tags that have content that have the same list id.
    """
    yield li
    w_namespace = get_namespace(li, 'w')
    current_numId = get_numId(li, w_namespace)
    starting_ilvl = get_ilvl(li, w_namespace)
    el = li
    while True:
        el = el.getnext()
        if el is None:
            break
        # If the tag has no content ignore it.
        if not has_text(el):
            continue

        # Stop the lists if you come across a list item that should be a
        # heading.
        if _is_top_level_upper_roman(el, meta_data):
            break

        if (
                is_li(el, meta_data) and
                (starting_ilvl > get_ilvl(el, w_namespace))):
            break

        new_numId = get_numId(el, w_namespace)
        if new_numId is None or new_numId == -1:
            # Not a p tag or a list item
            yield el
            continue
        # If the list id of the next tag is different that the previous that
        # means a new list being made (not nested)
        if current_numId != new_numId:
            # Not a subsequent list.
            break
        if is_last_li(el, meta_data, current_numId):
            yield el
            break
        yield el


@ensure_tag(['p'])
def get_ilvl(li, w_namespace):
    """
    The ilvl on an li tag tells the li tag at what level of indentation this
    tag is at. This is used to determine if the li tag needs to be nested or
    not.
    """
    ilvls = li.xpath('.//w:ilvl', namespaces=li.nsmap)
    if len(ilvls) == 0:
        return -1
    return int(ilvls[0].get('%sval' % w_namespace))


@ensure_tag(['p'])
def get_numId(li, w_namespace):
    """
    The numId on an li tag maps to the numbering dictionary along side the ilvl
    to determine what the list should look like (unordered, digits, lower
    alpha, etc)
    """
    numIds = li.xpath('.//w:numId', namespaces=li.nsmap)
    if len(numIds) == 0:
        return -1
    return numIds[0].get('%sval' % w_namespace)


def create_list(list_type):
    """
    Based on the passed in list_type create a list objects (ol/ul). In the
    future this function will also deal with what the numbering of an ordered
    list should look like.
    """
    list_types = {
        'bullet': 'ul',
    }
    el = etree.Element(list_types.get(list_type, 'ol'))
    # These are the supported list style types and their conversion to css.
    list_type_conversions = {
        'decimal': DEFAULT_LIST_NUMBERING_STYLE,
        'decimalZero': 'decimal-leading-zero',
        'upperRoman': 'upper-roman',
        'lowerRoman': 'lower-roman',
        'upperLetter': 'upper-alpha',
        'lowerLetter': 'lower-alpha',
        'ordinal': DEFAULT_LIST_NUMBERING_STYLE,
        'cardinalText': DEFAULT_LIST_NUMBERING_STYLE,
        'ordinalText': DEFAULT_LIST_NUMBERING_STYLE,
    }
    if list_type != 'bullet':
        el.set(
            'data-list-type',
            list_type_conversions.get(list_type, DEFAULT_LIST_NUMBERING_STYLE),
        )
    return el


@ensure_tag(['tc'])
def get_v_merge(tc):
    """
    vMerge is what docx uses to denote that a table cell is part of a rowspan.
    The first cell to have a vMerge is the start of the rowspan, and the vMerge
    will be denoted with 'restart'. If it is anything other than restart then
    it is a continuation of another rowspan.
    """
    if tc is None:
        return None
    v_merges = tc.xpath('.//w:vMerge', namespaces=tc.nsmap)
    if len(v_merges) != 1:
        return None
    v_merge = v_merges[0]
    return v_merge


@ensure_tag(['tc'])
def get_grid_span(tc):
    """
    gridSpan is what docx uses to denote that a table cell has a colspan. This
    is much more simple than rowspans in that there is a one-to-one mapping
    from gridSpan to colspan.
    """
    w_namespace = get_namespace(tc, 'w')
    grid_spans = tc.xpath('.//w:gridSpan', namespaces=tc.nsmap)
    if len(grid_spans) != 1:
        return 1
    grid_span = grid_spans[0]
    return int(grid_span.get('%sval' % w_namespace))


@ensure_tag(['tr'])
def get_td_at_index(tr, index):
    """
    When calculating the rowspan for a given cell it is required to find all
    table cells 'below' the initial cell with a v_merge. This function will
    return the td element at the passed in index, taking into account colspans.
    """
    current = 0
    for td in tr.xpath('.//w:tc', namespaces=tr.nsmap):
        if index == current:
            return td
        current += get_grid_span(td)


@ensure_tag(['tbl'])
def get_rowspan_data(table):
    w_namespace = get_namespace(table, 'w')

    # We need to keep track of what table row we are on as well as which table
    # cell we are on.
    tr_index = 0
    td_index = 0

    # Get a list of all the table rows.
    tr_rows = list(table.xpath('.//w:tr', namespaces=table.nsmap))

    # Loop through each table row.
    for tr in table.xpath('.//w:tr', namespaces=table.nsmap):
        # Loop through each table cell.
        for td in tr.xpath('.//w:tc', namespaces=tr.nsmap):
            # Check to see if this cell has a v_merge
            v_merge = get_v_merge(td)

            # If not increment the td_index and move on
            if v_merge is None:
                td_index += get_grid_span(td)
                continue

            # If it does have a v_merge we need to see if it is the ``root``
            # table cell (the first in a row to have a rowspan)
            # If the value is restart then this is the table cell that needs
            # the rowspan.
            if v_merge.get('%sval' % w_namespace) == 'restart':
                row_span = 1
                # Loop through each table row after the current one.
                for tr_el in tr_rows[tr_index + 1:]:
                    # Get the table cell at the current td_index.
                    td_el = get_td_at_index(tr_el, td_index)
                    td_el_v_merge = get_v_merge(td_el)

                    # If the td_ell does not have a v_merge then the rowspan is
                    # done.
                    if td_el_v_merge is None:
                        break
                    val = td_el_v_merge.get('%sval' % w_namespace)
                    # If the v_merge is restart then there is another cell that
                    # needs a rowspan, so the current cells rowspan is done.
                    if val == 'restart':
                        break
                    # Increment the row_span
                    row_span += 1
                yield row_span
            # Increment the indexes.
            td_index += get_grid_span(td)
        tr_index += 1
        # Reset the td_index when we finish each table row.
        td_index = 0


@ensure_tag(['b', 'i', 'u'])
def style_is_false(style):
    """
    For bold, italics and underline. Simply checking to see if the various tags
    are present will not suffice. If the tag is present and set to False then
    the style should not be present.
    """
    if style is None:
        return False
    w_namespace = get_namespace(style, 'w')
    return style.get('%sval' % w_namespace) != 'false'


@ensure_tag(['r'])
def is_bold(r):
    """
    The function will return True if the r tag passed in is considered bold.
    """
    w_namespace = get_namespace(r, 'w')
    rpr = r.find('%srPr' % w_namespace)
    if rpr is None:
        return False
    bold = rpr.find('%sb' % w_namespace)
    return style_is_false(bold)


@ensure_tag(['r'])
def is_italics(r):
    """
    The function will return True if the r tag passed in is considered
    italicized.
    """
    w_namespace = get_namespace(r, 'w')
    rpr = r.find('%srPr' % w_namespace)
    if rpr is None:
        return False
    italics = rpr.find('%si' % w_namespace)
    return style_is_false(italics)


@ensure_tag(['r'])
def is_underlined(r):
    """
    The function will return True if the r tag passed in is considered
    underlined.
    """
    w_namespace = get_namespace(r, 'w')
    rpr = r.find('%srPr' % w_namespace)
    if rpr is None:
        return False
    underline = rpr.find('%su' % w_namespace)
    return style_is_false(underline)


@ensure_tag(['p'])
def is_title(p):
    """
    Certain p tags are denoted as ``Title`` tags. This function will return
    True if the passed in p tag is considered a title.
    """
    w_namespace = get_namespace(p, 'w')
    styles = p.xpath('.//w:pStyle', namespaces=p.nsmap)
    if len(styles) == 0:
        return False
    style = styles[0]
    return style.get('%sval' % w_namespace) == 'Title'


@ensure_tag(['r'])
def get_text_run_content_data(r):
    """
    It turns out that r tags can contain both t tags and drawing tags. Since we
    need both, this function will return them in the order in which they are
    found.
    """
    w_namespace = get_namespace(r, 'w')
    valid_elements = (
        '%st' % w_namespace,
        '%sdrawing' % w_namespace,
        '%spict' % w_namespace,
        '%sbr' % w_namespace,
    )
    for el in r:
        if el.tag in valid_elements:
            yield el


@ensure_tag(['drawing', 'pict'])
def get_image_id(drawing):
    r_namespace = get_namespace(drawing, 'r')
    for el in drawing.iter():
        # For drawing
        image_id = el.get('%sembed' % r_namespace)
        if image_id is not None:
            return image_id
        # For pict
        if 'v' not in el.nsmap:
            continue
        v_namespace = get_namespace(drawing, 'v')
        if el.tag == '%simagedata' % v_namespace:
            image_id = el.get('%sid' % r_namespace)
            if image_id is not None:
                return image_id


@ensure_tag(['p'])
def whole_line_styled(p):
    """
    Checks to see if the whole p tag will end up being bold or italics. Returns
    a tuple (boolean, boolean). The first boolean will be True if the whole
    line is bold, False otherwise. The second boolean will be True if the whole
    line is italics, False otherwise.
    """
    r_tags = p.xpath('.//w:r', namespaces=p.nsmap)
    tags_are_bold = [
        is_bold(r) or is_underlined(r) for r in r_tags
    ]
    tags_are_italics = [
        is_italics(r) for r in r_tags
    ]
    return all(tags_are_bold), all(tags_are_italics)


MetaData = namedtuple(
    'MetaData',
    [
        'numbering_dict',
        'relationship_dict',
        'styles_dict',
        'font_sizes_dict',
        'image_handler',
        'image_sizes',
    ],
)


###
# Pre-processing
###


def get_numbering_info(tree):
    """
    There is a separate file called numbering.xml that stores how lists should
    look (unordered, digits, lower case letters, etc.). Parse that file and
    return a dictionary of what each combination should be based on list Id and
    level of indentation.
    """
    if tree is None:
        return {}
    w_namespace = get_namespace(tree, 'w')
    num_ids = {}
    result = defaultdict(dict)
    # First find all the list types
    for list_type in tree.findall('%snum' % w_namespace):
        list_id = list_type.get('%snumId' % w_namespace)

        # Each list type is assigned an abstractNumber that defines how lists
        # should look.
        abstract_number = list_type.find('%sabstractNumId' % w_namespace)
        num_ids[abstract_number.get('%sval' % w_namespace)] = list_id

    # Loop through all the abstractNumbers
    for abstract_number in tree.findall('%sabstractNum' % w_namespace):
        abstract_num_id = abstract_number.get('%sabstractNumId' % w_namespace)
        # If we find an abstractNumber that is not being used in the document
        # then ignore it.
        if abstract_num_id not in num_ids:
            continue

        # Get the level of the abstract number.
        for lvl in abstract_number.findall('%slvl' % w_namespace):
            ilvl = int(lvl.get('%silvl' % w_namespace))
            lvl_format = lvl.find('%snumFmt' % w_namespace)
            list_style = lvl_format.get('%sval' % w_namespace)
            # Based on the list type and the ilvl (indentation level) store the
            # needed style.
            result[num_ids[abstract_num_id]][ilvl] = list_style
    return result


def get_style_dict(tree):
    """
    Some things that are considered lists are actually supposed to be H tags
    (h1, h2, etc.) These can be denoted by their styleId
    """
    # This is a partial document and actual h1 is the document title, which
    # will be displayed elsewhere.
    headers = {
        'heading 1': 'h2',
        'heading 2': 'h3',
        'heading 3': 'h4',
        'heading 4': 'h5',
        'heading 5': 'h6',
        'heading 6': 'h6',
        'heading 7': 'h6',
        'heading 8': 'h6',
        'heading 9': 'h6',
        'heading 10': 'h6',
    }
    if tree is None:
        return {}
    w_namespace = get_namespace(tree, 'w')
    result = {}
    for el in tree:
        style_id = el.get('%sstyleId' % w_namespace)
        el_result = {
            'header': False,
            'font_size': None,
            'based_on': None,
        }
        # Get the header info
        name = el.find('%sname' % w_namespace)
        if name is None:
            continue
        value = name.get('%sval' % w_namespace).lower()
        if value in headers:
            el_result['header'] = headers[value]

        # Get the size info.
        rpr = el.find('%srPr' % w_namespace)
        if rpr is None:
            continue
        size = rpr.find('%ssz' % w_namespace)
        if size is None:
            el_result['font_size'] = None
        else:
            el_result['font_size'] = size.get('%sval' % w_namespace)

        # Get based on info.
        based_on = el.find('%sbasedOn' % w_namespace)
        if based_on is None:
            el_result['based_on'] = None
        else:
            el_result['based_on'] = based_on.get('%sval' % w_namespace)
        result[style_id] = el_result
    return result


def get_image_sizes(tree):
    drawings = []
    result = {}
    w_namespace = get_namespace(tree, 'w')
    for el in tree.iter():
        if el.tag == '%sdrawing' % w_namespace:
            drawings.append(el)
    for d in drawings:
        for el in d.iter():
            if 'a' not in el.nsmap:
                continue
            a_namespace = get_namespace(el, 'a')
            if el.tag == '%sxfrm' % a_namespace:
                ext = el.find('%sext' % a_namespace)
                cx = int(ext.get('cx')) / EMUS_PER_PIXEL
                cy = int(ext.get('cy')) / EMUS_PER_PIXEL
                result[get_image_id(d)] = (cx, cy)
    return result


def get_relationship_info(tree, media, image_sizes):
    """
    There is a separate file holds the targets to links as well as the targets
    for images. Return a dictionary based on the relationship id and the
    target.
    """
    if tree is None:
        return {}
    result = {}
    # Loop through each relationship.
    for el in tree.iter():
        el_id = el.get('Id')
        if el_id is None:
            continue
        # Store the target in the result dict.
        target = el.get('Target')
        if any(
                target.lower().endswith(ext) for
                ext in IMAGE_EXTENSIONS_TO_SKIP):
            continue
        if target in media:
            image_size = image_sizes.get(el_id)
            target = convert_image(media[target], image_size)
        # cgi will replace things like & < > with &amp; &lt; &gt;
        result[el_id] = cgi.escape(target)

    return result


def get_font_sizes_dict(tree, styles_dict):
    font_sizes_dict = defaultdict(int)
    # Get all the fonts sizes and how often they are used in a dict.
    for p in tree.xpath('//w:p', namespaces=tree.nsmap):
        # If this p tag is a natural header, skip it
        if is_natural_header(p, styles_dict):
            continue
        if _is_li(p):
            continue
        font_size = get_font_size(p, styles_dict)
        if font_size is None:
            continue
        font_sizes_dict[font_size] += 1

    # Find the most used font size.
    most_used_font_size = -1
    highest_count = -1
    for size, count in font_sizes_dict.items():
        if count > highest_count:
            highest_count = count
            most_used_font_size = size
    # Consider the most used font size to be the 'default' font size. Any font
    # size that is different will be considered an h tag.
    result = {}
    for size in font_sizes_dict:
        if size is None:
            continue
        if int(size) > int(most_used_font_size):
            # Not an h tag
            result[size] = 'h2'
        else:
            result[size] = None
    return result


def _get_document_data(f, image_handler=None):
    '''
    ``f`` is a ``ZipFile`` that is open
    Extract out the document data, numbering data and the relationship data.
    '''
    if image_handler is None:
        def image_handler(image_id, relationship_dict):
            return relationship_dict.get(image_id)

    document_xml = None
    numbering_xml = None
    relationship_xml = None
    styles_xml = None
    parser = etree.XMLParser(strip_cdata=False)
    path, _ = os.path.split(f.filename)
    media = {}
    image_sizes = {}
    # Loop through the files in the zip file.
    for item in f.infolist():
        # This file holds all the content of the document.
        if item.filename == 'word/document.xml':
            xml = f.read(item.filename)
            document_xml = etree.fromstring(xml, parser)
        # This file tells document.xml how lists should look.
        elif item.filename == 'word/numbering.xml':
            xml = f.read(item.filename)
            numbering_xml = etree.fromstring(xml, parser)
        elif item.filename == 'word/styles.xml':
            xml = f.read(item.filename)
            styles_xml = etree.fromstring(xml, parser)
        # This file holds the targets for hyperlinks and images.
        elif item.filename == 'word/_rels/document.xml.rels':
            xml = f.read(item.filename)
            try:
                relationship_xml = etree.fromstring(xml, parser)
            except XMLSyntaxError:
                relationship_xml = etree.fromstring('<xml></xml>', parser)
        if item.filename.startswith('word/media/'):
            # Strip off the leading word/
            media[item.filename[len('word/'):]] = f.extract(
                item.filename,
                path,
            )
    # Close the file pointer.
    f.close()

    # Get dictionaries for the numbering and the relationships.
    numbering_dict = get_numbering_info(numbering_xml)
    image_sizes = get_image_sizes(document_xml)
    relationship_dict = get_relationship_info(
        relationship_xml,
        media,
        image_sizes
    )
    styles_dict = get_style_dict(styles_xml)
    font_sizes_dict = defaultdict(int)
    if DETECT_FONT_SIZE:
        font_sizes_dict = get_font_sizes_dict(document_xml, styles_dict)
    meta_data = MetaData(
        numbering_dict=numbering_dict,
        relationship_dict=relationship_dict,
        styles_dict=styles_dict,
        font_sizes_dict=font_sizes_dict,
        image_handler=image_handler,
        image_sizes=image_sizes,
    )
    return document_xml, meta_data


###
# HTML Building functions
###


def get_ordered_list_type(meta_data, numId, ilvl):
    """
    Return the list type. If numId or ilvl not in the numbering dict then
    default to returning decimal.

    This function only cares about ordered lists, unordered lists get dealt
    with elsewhere.
    """

    # Early return if numId or ilvl are not valid
    numbering_dict = meta_data.numbering_dict
    if numId not in numbering_dict:
        return DEFAULT_LIST_NUMBERING_STYLE
    if ilvl not in numbering_dict[numId]:
        return DEFAULT_LIST_NUMBERING_STYLE
    return meta_data.numbering_dict[numId][ilvl]


def build_list(li_nodes, meta_data):
    """
    Build the list structure and return the root list
    """
    # Need to keep track of all incomplete nested lists.
    ol_dict = {}

    # Need to keep track of the current indentation level.
    current_ilvl = -1

    # Need to keep track of the current list id.
    current_numId = -1

    # Need to keep track of list that new li tags should be added too.
    current_ol = None

    # Store the first list created (the root list) for the return value.
    root_ol = None
    visited_nodes = []
    list_contents = []

    def _build_li(list_contents):
        data = '<br />'.join(t for t in list_contents if t is not None)
        return etree.XML('<li>%s</li>' % data)

    def _build_non_li_content(el, meta_data):
        w_namespace = get_namespace(el, 'w')
        if el.tag == '%stbl' % w_namespace:
            new_el, visited_nodes = build_table(el, meta_data)
            return etree.tostring(new_el), visited_nodes
        elif el.tag == '%sp' % w_namespace:
            return get_element_content(el, meta_data), [el]
        if has_text(el):
            raise UnintendedTag('Did not expect %s' % el.tag)

    def _merge_lists(ilvl, current_ilvl, ol_dict, current_ol):
        for i in reversed(range(ilvl, current_ilvl)):
            # Any list that is more indented that ilvl needs to
            # be merged to the list before it.
            if i not in ol_dict:
                continue
            if ol_dict[i] is not current_ol:
                if ol_dict[i] is current_ol:
                    continue
                ol_dict[i][-1].append(current_ol)
                current_ol = ol_dict[i]

        # Clean up finished nested lists.
        for key in list(ol_dict):
            if key > ilvl:
                del ol_dict[key]
        return current_ol

    for li_node in li_nodes:
        w_namespace = get_namespace(li_node, 'w')
        if not is_li(li_node, meta_data):
            # Get the content and visited nodes
            new_el, el_visited_nodes = _build_non_li_content(
                li_node,
                meta_data,
            )
            list_contents.append(new_el)
            visited_nodes.extend(el_visited_nodes)
            continue
        if list_contents:
            li_el = _build_li(list_contents)
            list_contents = []
            current_ol.append(li_el)
        # Get the data needed to build the current list item
        list_contents.append(get_element_content(
            li_node,
            meta_data,
        ))
        ilvl = get_ilvl(li_node, w_namespace)
        numId = get_numId(li_node, w_namespace)
        list_type = get_ordered_list_type(meta_data, numId, ilvl)

        # If the ilvl is greater than the current_ilvl or the list id is
        # changing then we have the first li tag in a nested list. We need to
        # create a new list object and update all of our variables for keeping
        # track.
        if (ilvl > current_ilvl) or (numId != current_numId):
            # Only create a new list
            ol_dict[ilvl] = create_list(list_type)
            current_ol = ol_dict[ilvl]
            current_ilvl = ilvl
            current_numId = numId
        # Both cases above are not True then we need to close all lists greater
        # than ilvl and then remove them from the ol_dict
        else:
            # Merge any nested lists that need to be merged.
            current_ol = _merge_lists(
                ilvl=ilvl,
                current_ilvl=current_ilvl,
                ol_dict=ol_dict,
                current_ol=current_ol,
            )

        # Set the root list after the first list is created.
        if root_ol is None:
            root_ol = current_ol

        # Set the current list.
        if ilvl in ol_dict:
            current_ol = ol_dict[ilvl]
        else:
            # In some instances the ilvl is not in the ol_dict, if that is the
            # case, create it here (not sure how this happens but it has
            # before.) Only do this if the current_ol is not the root_ol,
            # otherwise etree will crash.

            if current_ol is not root_ol:

                # Merge the current_ol into the root_ol. _merge_lists is not
                # equipped to handle this situation since the only way to get
                # into this block of code is to have mangled ilvls.
                root_ol[-1].append(current_ol)

                # Reset the current_ol
                current_ol = create_list(list_type)

        # Create the li element.
        visited_nodes.extend(list(li_node.iter()))

    # If a list item is the last thing in a document, then you will need to add
    # it here. Should probably figure out how to get the above logic to deal
    # with it.
    if list_contents:
        li_el = _build_li(list_contents)
        list_contents = []
        current_ol.append(li_el)

    # Merge up any nested lists that have not been merged.
    current_ol = _merge_lists(
        ilvl=0,
        current_ilvl=current_ilvl,
        ol_dict=ol_dict,
        current_ol=current_ol,
    )

    return root_ol, visited_nodes


@ensure_tag(['tr'])
def build_tr(tr, meta_data, row_spans):
    """
    This will return a single tr element, with all tds already populated.
    """

    # Create a blank tr element.
    tr_el = etree.Element('tr')
    w_namespace = get_namespace(tr, 'w')
    visited_nodes = []
    for el in tr:
        if el in visited_nodes:
            continue
        visited_nodes.append(el)
        # Find the table cells.
        if el.tag == '%stc' % w_namespace:
            v_merge = get_v_merge(el)
            # If there is a v_merge and it is not restart then this cell can be
            # ignored.
            if (
                    v_merge is not None and
                    v_merge.get('%sval' % w_namespace) != 'restart'):
                continue

            # Loop through each and build a list of all the content.
            texts = []
            for td_content in el:
                # Since we are doing look-a-heads in this loop we need to check
                # again to see if we have already visited the node.
                if td_content in visited_nodes:
                    continue

                # Check to see if it is a list or a regular paragraph.
                if is_li(td_content, meta_data):
                    # If it is a list, create the list and update
                    # visited_nodes.
                    li_nodes = get_single_list_nodes_data(
                        td_content,
                        meta_data,
                    )
                    list_el, list_visited_nodes = build_list(
                        li_nodes,
                        meta_data,
                    )
                    visited_nodes.extend(list_visited_nodes)
                    texts.append(etree.tostring(list_el))
                elif td_content.tag == '%stbl' % w_namespace:
                    table_el, table_visited_nodes = build_table(
                        td_content,
                        meta_data,
                    )
                    visited_nodes.extend(table_visited_nodes)
                    texts.append(etree.tostring(table_el))
                elif td_content.tag == '%stcPr' % w_namespace:
                    # Do nothing
                    visited_nodes.append(td_content)
                    continue
                else:
                    text = get_element_content(
                        td_content,
                        meta_data,
                        is_td=True,
                    )
                    texts.append(text)

            data = '<br />'.join(t for t in texts if t is not None)
            td_el = etree.XML('<td>%s</td>' % data)
            # if there is a colspan then set it here.
            colspan = get_grid_span(el)
            if colspan > 1:
                td_el.set('colspan', '%d' % colspan)
            v_merge = get_v_merge(el)

            # If this td has a v_merge and it is restart then set the rowspan
            # here.
            if (
                    v_merge is not None and
                    v_merge.get('%sval' % w_namespace) == 'restart'):
                rowspan = next(row_spans)
                td_el.set('rowspan', '%d' % rowspan)

            tr_el.append(td_el)
    return tr_el


@ensure_tag(['tbl'])
def build_table(table, meta_data):
    """
    This returns a table object with all rows and cells correctly populated.
    """

    # Create a blank table element.
    table_el = etree.Element('table')
    w_namespace = get_namespace(table, 'w')

    # Get the rowspan values for cells that have a rowspan.
    row_spans = get_rowspan_data(table)
    for el in table:
        if el.tag == '%str' % w_namespace:
            # Create the tr element.
            tr_el = build_tr(
                el,
                meta_data,
                row_spans,
            )
            # And append it to the table.
            table_el.append(tr_el)

    visited_nodes = list(table.iter())
    return table_el, visited_nodes


@ensure_tag(['t'])
def get_t_tag_content(
        t, parent, remove_bold, remove_italics, meta_data):
    """
    Generate the string data that for this particular t tag.
    """
    if t is None or t.text is None:
        return ''

    # Need to escape the text so that we do not accidentally put in text
    # that is not valid XML.
    # cgi will replace things like & < > with &amp; &lt; &gt;
    text = cgi.escape(t.text)

    # Wrap the text with any modifiers it might have (bold, italics or
    # underline)
    el_is_bold = not remove_bold and (
        is_bold(parent) or
        is_underlined(parent)
    )
    el_is_italics = not remove_italics and is_italics(parent)
    if el_is_bold:
        text = '<strong>%s</strong>' % text
    if el_is_italics:
        text = '<em>%s</em>' % text
    return text


def _get_image_size_from_image(target):
    image = Image.open(target)
    return image.size


def build_hyperlink(el, meta_data):
    # If we have a hyperlink we need to get relationship_id
    r_namespace = get_namespace(el, 'r')
    hyperlink_id = el.get('%sid' % r_namespace)

    # Once we have the hyperlink_id then we need to replace the
    # hyperlink tag with its child run tags.
    content = get_element_content(
        el,
        meta_data,
        remove_bold=True,
        remove_italics=True,
    )
    if not content:
        return ''
    if hyperlink_id in meta_data.relationship_dict:
        href = meta_data.relationship_dict[hyperlink_id]
        # Do not do any styling on hyperlinks
        return '<a href="%s">%s</a>' % (href, content)
    return ''


def build_image(el, meta_data):
    image_id = get_image_id(el)
    if image_id not in meta_data.relationship_dict:
        # This image does not have an image_id
        return ''
    src = meta_data.image_handler(
        image_id,
        meta_data.relationship_dict,
    )
    if image_id in meta_data.image_sizes:
        width, height = meta_data.image_sizes[image_id]
    else:
        target = meta_data.relationship_dict[image_id]
        width, height = _get_image_size_from_image(target)
    # Make sure the width and height are not zero
    if all((width, height)):
        return '<img src="%s" height="%d" width="%d" />' % (
            src,
            height,
            width,
        )
    else:
        return '<img src="%s" />' % src
    return ''


def get_text_run_content(el, meta_data, remove_bold, remove_italics):
    w_namespace = get_namespace(el, 'w')
    text_output = ''
    for child in get_text_run_content_data(el):
        if child.tag == '%st' % w_namespace:
            text_output += get_t_tag_content(
                child,
                el,
                remove_bold,
                remove_italics,
                meta_data,
            )
        elif child.tag == '%sbr' % w_namespace:
            text_output += '<br />'
        elif child.tag in (
                '%spict' % w_namespace,
                '%sdrawing' % w_namespace,
        ):
            text_output += build_image(child, meta_data)
        else:
            raise SyntaxNotSupported(
                '"%s" is not a supported content-containing '
                'text run child.' % child.tag
            )
    return text_output


@ensure_tag(['p', 'ins', 'smartTag', 'hyperlink'])
def get_element_content(
        p,
        meta_data,
        is_td=False,
        remove_italics=False,
        remove_bold=False,
):
    """
    P tags are made up of several runs (r tags) of text. This function takes a
    p tag and constructs the text that should be part of the p tag.

    image_handler should be a callable that returns the desired ``src``
    attribute for a given image.
    """

    # Only remove bold or italics if this tag is an h tag.
    # Td elements have the same look and feel as p/h elements. Right now we are
    # never putting h tags in td elements, as such if we are in a td we will
    # never be stripping bold/italics since that is only done on h tags
    if not is_td and is_header(p, meta_data):
        # Check to see if the whole line is bold or italics.
        remove_bold, remove_italics = whole_line_styled(p)

    p_text = ''
    w_namespace = get_namespace(p, 'w')
    if len(p) == 0:
        return ''
    # Only these tags contain text that we care about (eg. We don't care about
    # delete tags)
    content_tags = (
        '%sr' % w_namespace,
        '%shyperlink' % w_namespace,
        '%sins' % w_namespace,
        '%ssmartTag' % w_namespace,
    )
    elements_with_content = []
    for child in p:
        if child is None:
            break
        if child.tag in content_tags:
            elements_with_content.append(child)

    # Gather the content from all of the children
    for el in elements_with_content:
        # Hyperlinks and insert tags need to be handled differently than
        # r and smart tags.
        if el.tag in ('%sins' % w_namespace, '%ssmartTag' % w_namespace):
            p_text += get_element_content(
                el,
                meta_data,
                remove_bold=remove_bold,
                remove_italics=remove_italics,
            )
        elif el.tag == '%shyperlink' % w_namespace:
            p_text += build_hyperlink(el, meta_data)
        elif el.tag == '%sr' % w_namespace:
            p_text += get_text_run_content(
                el,
                meta_data,
                remove_bold=remove_bold,
                remove_italics=remove_italics,
            )
        else:
            raise SyntaxNotSupported(
                'Content element "%s" not handled.' % el.tag
            )

    # This function does not return a p tag since other tag types need this as
    # well (td, li).
    return p_text


def _strip_tag(tree, tag):
    """
    Remove all tags that have the tag name ``tag``
    """
    for el in tree.iter():
        if el.tag == tag:
            el.getparent().remove(el)


def get_zip_file_handler(file_path):
    return ZipFile(file_path)


def read_html_file(file_path):
    with open(file_path) as f:
        html = f.read()
    return html


def convert(file_path, image_handler=None, fall_back=None, converter=None):
    """
    ``file_path`` is a path to the file on the file system that you want to be
        converted to html.
    ``image_handler`` is a function that takes an image_id and a
        relationship_dict to generate the src attribute for images. (see readme
        for more details)
    ``fall_back`` is a function that takes a ``file_path``. This function will
        only be called if for whatever reason the conversion fails.
    ``converter`` is a function to convert a document that is not docx to docx
        (examples in docx2html.converters)

    Returns html extracted from ``file_path``
    """
    file_base, extension = os.path.splitext(os.path.basename(file_path))

    if extension == '.html' or extension == '.htm':
        return read_html_file(file_path)

    # Create the converted file as a file in the same dir with the
    # same name only with a .docx extension
    docx_path = replace_ext(file_path, '.docx')
    if extension == '.docx':
        # If the file is already html, just leave it in place.
        docx_path = file_path
    else:
        if converter is None:
            raise FileNotDocx('The file passed in is not a docx.')
        converter(docx_path, file_path)
        if not os.path.isfile(docx_path):
            if fall_back is None:
                raise ConversionFailed('Conversion to docx failed.')
            else:
                return fall_back(file_path)

    try:
        # Docx files are actually just zip files.
        zf = get_zip_file_handler(docx_path)
    except BadZipfile:
        raise MalformedDocx('This file is not a docx')

    # Need to populate the xml based on word/document.xml
    tree, meta_data = _get_document_data(zf, image_handler)
    return create_html(tree, meta_data)


def create_html(tree, meta_data):

    # Start the return value
    new_html = etree.Element('html')

    w_namespace = get_namespace(tree, 'w')
    visited_nodes = []

    _strip_tag(tree, '%ssectPr' % w_namespace)
    for el in tree.iter():
        # The way lists are handled could double visit certain elements; keep
        # track of which elements have been visited and skip any that have been
        # visited already.
        if el in visited_nodes:
            continue
        header_value = is_header(el, meta_data)
        if is_header(el, meta_data):
            p_text = get_element_content(el, meta_data)
            if p_text == '':
                continue
            new_html.append(
                etree.XML('<%s>%s</%s>' % (
                    header_value,
                    p_text,
                    header_value,
                ))
            )
        elif el.tag == '%sp' % w_namespace:
            # Strip out titles.
            if is_title(el):
                continue
            if is_li(el, meta_data):
                # Parse out the needed info from the node.
                li_nodes = get_single_list_nodes_data(el, meta_data)
                new_el, list_visited_nodes = build_list(
                    li_nodes,
                    meta_data,
                )
                visited_nodes.extend(list_visited_nodes)
            # Handle generic p tag here.
            else:
                p_text = get_element_content(el, meta_data)
                # If there is not text do not add an empty tag.
                if p_text == '':
                    continue

                new_el = etree.XML('<p>%s</p>' % p_text)
            new_html.append(new_el)

        elif el.tag == '%stbl' % w_namespace:
            table_el, table_visited_nodes = build_table(
                el,
                meta_data,
            )
            visited_nodes.extend(table_visited_nodes)
            new_html.append(table_el)
            continue

        # Keep track of visited_nodes
        visited_nodes.append(el)
    result = etree.tostring(
        new_html,
        method='html',
        with_tail=True,
    )
    return _make_void_elements_self_close(result)


def _make_void_elements_self_close(html):
    #XXX Hack not sure how to get etree to do this by default.
    void_tags = [
        r'br',
        r'img',
    ]
    for tag in void_tags:
        regex = re.compile(r'<%s.*?>' % tag)
        matches = regex.findall(html)
        for match in matches:
            new_tag = match.strip('<>')
            new_tag = '<%s />' % new_tag
            html = re.sub(match, new_tag, html)
    return html
