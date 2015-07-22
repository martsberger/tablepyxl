# This is where we handle translating css styles into openpyxl styles
# and cascading those from parent to child in the dom.

from openpyxl.styles import Font, Alignment, PatternFill, Style


def style_string_to_dict(style):
    """
    Convert css style string to a python dictionary
    """
    def clean_split(string, delim):
        return (s.strip() for s in string.split(delim))
    styles = [clean_split(s, ":") for s in style.split(";") if ":" in s]
    return dict(styles)


def style_dict_to_Style(style):
    """
    change css style (stored in a python dictionary) to openpyxl Style
    """
    # Font
    font_kwargs = {'bold': style.get('font-weight') == 'bold',
                   'color': style.get('color', '000000')}
    font = Font(**font_kwargs)

    # Alignment
    alignment_kwargs = {'horizontal': style.get('text-align', 'general'),
                        'wrap_text': style.get('white-space', 'nowrap') == 'normal'}
    alignment = Alignment(**alignment_kwargs)

    # Fill
    fill = PatternFill()

    pyxl_style = Style(font=font, alignment=alignment, fill=fill)

    return pyxl_style


class StyleDict(dict):
    """
    It's like a dictionary, but it looks for items in the parent dictionary
    """
    def __init__(self, *args, **kwargs):
        self.parent = kwargs.pop('parent', None)
        super(StyleDict, self).__init__(*args, **kwargs)

    def __getitem__(self, item):
        if item in self:
            return super(StyleDict, self).__getitem__(item)
        elif self.parent:
            return self.parent[item]
        else:
            raise KeyError('%s not found' % item)

    def get(self, k, d=None):
        try:
            return self[k]
        except KeyError:
            return d


class Element(object):
    """
    Our base class for representing an html element along with a cascading style.
    The element is created along with a parent so that the StyleDict that we store
    can point to the parent's StyleDict.
    """
    def __init__(self, element, parent=None):
        self.element = element
        parent_style = parent.style_dict if parent else None
        self.style_dict = StyleDict(style_string_to_dict(element.get('style', '')), parent=parent_style)
        self._style_cache = None

    # TODO: This method is probably not necessary since we implemented StyleDict
    def get_style(self, key):
        return self.style_dict.get(key)

    def style(self):
        """
        Turn the css styles for this element into an openpyxl Style
        """
        if not self._style_cache:
            self._style_cache = style_dict_to_Style(self.style_dict)
        return self._style_cache


# The concrete implementations of Elements are semantically named for
# the types of elements we are interested in. This defines a very
# concrete tree structure for html tables that we expect to deal with.
# I prefer this compared to allowing Element to have an abitrary number
# of children and dealing with an abstract element tree.
class Table(Element):
    def __init__(self, table):
        """
        takes an html table object (from BeautifulSoup)
        """
        super(Table, self).__init__(table)
        self.head = TableHead(table.thead, parent=self) if table.thead else None
        self.body = TableBody(table.tbody or table, parent=self)


class TableHead(Element):
    def __init__(self, head, parent=None):
        super(TableHead, self).__init__(head, parent=parent)
        self.rows = [TableRow(tr, parent=self) for tr in head.find_all('tr')]


class TableBody(Element):
    def __init__(self, body, parent=None):
        super(TableBody, self).__init__(body, parent=parent)
        self.rows = [TableRow(tr, parent=self) for tr in body.find_all('tr')]


class TableRow(Element):
    def __init__(self, tr, parent=None):
        super(TableRow, self).__init__(tr, parent=parent)
        self.cells = [TableCell(cell, parent=self) for cell in tr.find_all('th') + tr.find_all('td')]


class TableCell(Element):
    pass
