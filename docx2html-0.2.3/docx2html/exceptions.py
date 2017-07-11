class Docx2HtmlException(Exception):
    pass


class ConversionFailed(Docx2HtmlException):
    pass


class FileNotDocx(Docx2HtmlException):
    pass


class MalformedDocx(Docx2HtmlException):
    pass


class UnintendedTag(Docx2HtmlException):
    pass


class SyntaxNotSupported(Docx2HtmlException):
    pass
