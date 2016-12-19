# This is where we handle translating css styles into openpyxl styles
# and cascading those from parent to child in the dom.

from openpyxl.cell import Cell
from openpyxl.styles import Font, Alignment, PatternFill, NamedStyle, Border, Side, Color
from openpyxl.styles.fills import FILL_SOLID
from openpyxl.styles.numbers import FORMAT_CURRENCY_USD_SIMPLE
from openpyxl.styles.colors import BLACK

FORMAT_DATE_MMDDYYYY = 'mm/dd/yyyy'


def colormap(color):
    """
    Convenience for looking up known colors
    """
    cmap = {'black': BLACK}
    return cmap.get(color, color)


def style_string_to_dict(style):
    """
    Convert css style string to a python dictionary
    """
    def clean_split(string, delim):
        return (s.strip() for s in string.split(delim))
    styles = [clean_split(s, ":") for s in style.split(";") if ":" in s]
    return dict(styles)


def get_side(style, name):
    return {'border_style': style.get('border-{}-style'.format(name)),
            'color': colormap(style.get('border-{}-color'.format(name)))}

known_styles = {}


def style_dict_to_Style(style):
    """
    change css style (stored in a python dictionary) to openpyxl Style
    """
    if style not in known_styles:
        # Font
        font = Font(bold=style.get('font-weight') == 'bold',
                    color=style.get_color('color', '000000'),
                    size=style.get('font-size'))

        # Alignment
        alignment = Alignment(horizontal=style.get('text-align', 'general'),
                              vertical=style.get('vertical-align'),
                              wrap_text=style.get('white-space', 'nowrap') == 'normal')

        # Fill
        bg_color = style.get_color('background-color')
        fg_color = style.get_color('foreground-color', Color())
        fill_type = style.get('fill-type')
        if bg_color:
            fill = PatternFill(fill_type=fill_type or FILL_SOLID,
                               start_color=bg_color,
                               end_color=fg_color)
        else:
            fill = PatternFill()

        # Border
        border = Border(left=Side(**get_side(style, 'left')),
                        right=Side(**get_side(style, 'right')),
                        top=Side(**get_side(style, 'top')),
                        bottom=Side(**get_side(style, 'bottom')),
                        diagonal=Side(**get_side(style, 'diagonal')),
                        diagonal_direction=None,
                        outline=Side(**get_side(style, 'outline')),
                        vertical=None,
                        horizontal=None)

        pyxl_style = NamedStyle(font=font, fill=fill, alignment=alignment, border=border)

        known_styles[style] = pyxl_style

    return known_styles[style]


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

    def __hash__(self):
        return hash(tuple([(k, self.get(k)) for k in self._keys()]))

    # Yielding the keys avoids creating unnecessary data structures
    # and happily works with both python2 and python3 where the
    # .keys() method is a dictionary_view in python3 and a list in python2.
    def _keys(self):
        yielded = set()
        for k in self.keys():
            yielded.add(k)
            yield k
        if self.parent:
            for k in self.parent._keys():
                if k not in yielded:
                    yielded.add(k)
                    yield k

    def get(self, k, d=None):
        try:
            return self[k]
        except KeyError:
            return d

    def get_color(self, k, d=None):
        """
        Strip leading # off colors if necessary
        """
        color = self.get(k, d)
        if hasattr(color, 'startswith') and color.startswith('#'):
            return color[1:]
        return color


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
    CELL_TYPES = {'TYPE_STRING', 'TYPE_FORMULA', 'TYPE_NUMERIC', 'TYPE_BOOL', 'TYPE_CURRENCY',
                  'TYPE_NULL', 'TYPE_INLINE', 'TYPE_ERROR', 'TYPE_FORMULA_CACHE_STRING'}

    def __init__(self, *args, **kwargs):
        super(TableCell, self).__init__(*args, **kwargs)
        self.value = self.element.get_text(separator="\n", strip=True)

    def data_type(self):
        cell_type = self.CELL_TYPES & set(self.element.get('class', []))
        if cell_type:
            cell_type = cell_type.pop()
            if cell_type == 'TYPE_CURRENCY':
                cell_type = 'TYPE_NUMERIC'
        else:
            cell_type = 'TYPE_STRING'
        return getattr(Cell, cell_type)

    def number_format(self):
        if 'TYPE_CURRENCY' in self.element.get('class', []):
            return FORMAT_CURRENCY_USD_SIMPLE
        if 'TYPE_DATE' in self.element.get('class', []):
            return FORMAT_DATE_MMDDYYYY
        if self.data_type() == Cell.TYPE_NUMERIC:
            try:
                int(self.value)
            except ValueError:
                return '#,##0.##'
            else:
                return '#,##0'

    def format(self, cell):
        style = self.style()
        cell.font = style.font
        cell.alignment = style.alignment
        cell.fill = style.fill
        cell.border = style.border
        data_type = self.data_type()
        if data_type:
            cell.data_type = data_type
        number_format = self.number_format()
        if number_format:
            cell.number_format = number_format
