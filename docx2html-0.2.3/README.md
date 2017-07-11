=========
docx2html
=========

Convert a docx (OOXML) file to semantic HTML.
All of Word formatting nonsense is stripped away and
you're left with a cleanly-formatted version of the content.


Usage
=====

    >>> from docx2html import convert
    >>> html = convert('path/to/docx/file')


Running Tests for Development
=============================

::

     $ virtualenv path/to/new/virtualenv
     $ source path/to/new/virtualenv/bin/activate
     $ cd path/to/workspace
     $ git clone git://github.com/PolicyStat/docx2html.git
     $ cd docx2html
     $ pip install .
     $ pip install -r test_requirements.txt
     $ ./run_tests.sh

Description
===========

docx2html is designed to take a docx file and extract the content out and
convert that content to html. It does not care about styles or fonts or
anything that changes how the content is displayed (with few exceptions). Below
is a list of what currently works:

* Paragraphs
    * Bold
    * Italics
    * Underline
    * Hyperlinks
* Lists
    * Nested lists
    * List styles (letters, roman numerals, etc.)
    * Tables
    * Paragraphs
* Tables
    * Rowspans
    * Colspans
    * Nested tables
    * Lists
* Images
    * Resizing
    * Converting to smaller formats (for bitmaps and tiffs)
    * There is a hook to allow setting the src of the image tag out of context,
      more on this later
* Headings
    * Simple headings
    * Root level lists that are upper case roman numerals get converted to h2
      tags

Handling embedded images
------------------------

docx2html allows you to specify how you would like to handle image uploading.
For example, you might be uploading your images to Amazon S3 eg:
Note: This documentation sucks, so you might need to read the source.

::

    import os.path
    from shutil import copyfile

    from docx2html import convert

    def handle_image(image_id, relationship_dict):
        image_path = relationship_dict[image_id]
        # Now do something to the image. Let's move it somewhere.
        _, filename = os.path.split(image_path)
        destination_path = os.path.join('/tmp', filename)
        copyfile(image_path, destination_path)

        # Return the `src` attribute to be used in the img tag
        return 'file://%s' % destination

    html = convert('path/to/docx/file', image_handler=handle_image)

Naming Conventions
------------------

There are two main naming conventions in the source for docx2html there are
*build* functions, which will return an etree element that represents HTML. And
there are *get_content* functions which return string representations of HTML.
