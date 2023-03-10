from openpyxl.styles import (Alignment, Border, Font, NamedStyle, PatternFill,
                             Protection, Side)
from titles import Titles


class Styles:

    _title_font = Font(
        name='Calibri',
        size=12,
        bold=True,
        italic=False,
        vertAlign=None,
        underline='none',
        strike=False,
        color='FF0000'
    )

    _header_font = Font(
        name='Calibri',
        size=11,
        bold=True,
        italic=False,
        vertAlign=None,
        underline='none',
        strike=False,
        color='000000'
    )

    _base_font = Font(
        name='Calibri',
        size=11,
        bold=False,
        italic=False,
        vertAlign=None,
        underline='none',
        strike=False,
        color='000000'
    )

    _footer_font = Font(
        name='Calibri',
        size=11,
        bold=True,
        italic=False,
        vertAlign=None,
        underline='none',
        strike=False,
        color='000000'
    )

    _header_fill = PatternFill(
        fill_type='solid',
        fgColor="ADD8E6",
        # start_color='FFFFFFFF',
        # end_color='FF000000'
    )

    _base_border = Border(
        left=Side(border_style='thin', color='000000'),
        right=Side(border_style='thin', color='000000'),
        top=Side(border_style='thin', color='000000'),
        bottom=Side(border_style='thin', color='000000'),
        # diagonal=Side(border_style=None, color='FF000000'),
        # diagonal_direction=0,
        # outline=Side(border_style=None, color='FF000000'),
        # vertical=Side(border_style=None, color='FF000000'),
        # horizontal=Side(border_style=None, color='FF000000')
    )

    _base_alignment = Alignment(
        horizontal='center',
        vertical='center',
        text_rotation=0,
        wrap_text=True,
        shrink_to_fit=False,
        indent=0
    )

    _direction_alignment = Alignment(
        horizontal='left',
        vertical='center',
        text_rotation=0,
        wrap_text=True,
        shrink_to_fit=False,
        indent=0
    )

    _number_format = 'General'

    _protection = Protection(
        locked=True,
        hidden=False
    )

    @classmethod
    @property
    def style_1(cls):
        base_style = NamedStyle(name=Titles.base_style)
        base_style.font = cls._base_font
        base_style.border = cls._base_border
        base_style.alignment = cls._base_alignment
        return base_style

    @classmethod
    @property
    def style_2(cls):
        footer_style = NamedStyle(name=Titles.footer_style)
        footer_style.font = cls._footer_font
        footer_style.border = cls._base_border
        footer_style.alignment = cls._base_alignment
        return footer_style

    @classmethod
    @property
    def style_3(cls):
        header_style = NamedStyle(name=Titles.header_style)
        header_style.font = cls._header_font
        header_style.border = cls._base_border
        header_style.alignment = cls._base_alignment
        header_style.fill = cls._header_fill
        return header_style

    @classmethod
    @property
    def style_4(cls):
        title_style = NamedStyle(name=Titles.title_style)
        title_style.font = cls._title_font
        return title_style

    @classmethod
    @property
    def style_5(cls):
        direction_style = NamedStyle(name=Titles.direction_style)
        direction_style.font = cls._header_font
        direction_style.border = cls._base_border
        direction_style.alignment = cls._direction_alignment
        return direction_style
