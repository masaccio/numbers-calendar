import argparse
import calendar
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from datetime import date
from locale import getlocale
from sys import exit

import pycountry
from dateutil.relativedelta import relativedelta
from holidays import HolidayBase, country_holidays, list_supported_countries
from numbers_parser import Border, Document, xl_range
from xlsxwriter import Workbook

from spreadsheet_calendar import __version__


def generate_month_map():
    month_map = {}
    for mon in range(1, 13):
        month_map[calendar.month_name[mon]] = mon
        month_map[calendar.month_abbr[mon]] = mon
    return month_map


def generate_weekday_map():
    weekday_map = {}
    for day in range(0, 7):
        weekday_map[calendar.day_name[day]] = day + 1
        weekday_map[calendar.day_abbr[day]] = day + 1
    return weekday_map


DEFAULT_LOCALE = getlocale()[0][-2:]
DEFAULT_WEEKENDS = [6, 7]
DEFAULT_YEAR = date.today().year
MONTH_MAP = generate_month_map()
WEEKDAY_MAP = generate_weekday_map()


def valid_month(month):
    """Retuns a month int for Argparse if valid or raises an exception"""
    try:
        _ = MONTH_MAP[month]
    except KeyError:
        raise argparse.ArgumentTypeError(f"'{month}' is not a valid month abbreviation or name")
    return MONTH_MAP[month]


def valid_weekday(weekday):
    """Retuns a weekday int for Argparse if valid or raises an exception"""
    try:
        _ = WEEKDAY_MAP[weekday]
    except KeyError:
        raise argparse.ArgumentTypeError(f"'{weekday}' is not a valid day abbreviation or name")
    return WEEKDAY_MAP[weekday]


def valid_year(year):
    """Retuns a year int for Argparse if valid or raises an exception"""
    try:
        # Also raises on invalid int conversion
        if int(year) < 0:
            raise ValueError()
    except ValueError:
        raise argparse.ArgumentTypeError(f"'{year}' is not a valid year")
    return int(year)


def generate_country_lookups():
    alpha2_to_country = {}
    country_to_alpha2 = {}
    alpha2_regions = {}
    for alpha2, subdivs in list_supported_countries().items():
        country_data = pycountry.countries.get(alpha_2=alpha2)
        if country_data is not None:
            alpha2_to_country[alpha2] = country_data.name
            country_to_alpha2[country_data.name] = alpha2
            alpha2_regions[alpha2] = subdivs

    return alpha2_to_country, country_to_alpha2, alpha2_regions


def command_line_parser():
    parser = argparse.ArgumentParser(description="Create spreadsheet calendars using python")
    parser.add_argument("-V", "--version", action="store_true")
    parser.add_argument(
        "--list-countries", action="store_true", help="List available country code and exit"
    )
    parser.add_argument(
        "--list-regions",
        action="store_true",
        help="List available regions in the selected country and exit",
    )
    parser.add_argument(
        "--no-holiday-weekends",
        action="store_true",
        default=False,
        help="Don't color weekends with holidays",
    )
    parser.add_argument(
        "--format",
        default="excel",
        choices=["numbers", "excel"],
        help="spreadsheet output format",
    )
    parser.add_argument(
        "--start-month",
        type=valid_month,
        metavar="month",
        default="Jan",
        help="Start month for calendar (default: Jan)",
    )
    parser.add_argument(
        "--weekend",
        type=valid_weekday,
        default=DEFAULT_WEEKENDS,
        nargs="*",
        action="extend",
        metavar="day",
        help="Days to highlight as weekends (default: Sat, Sun)",
    )
    parser.add_argument(
        "-o",
        "--output",
        metavar="filename",
        help="Output file (default: calendar.numbers/calendar.xlsx)",
    )
    parser.add_argument(
        "--country",
        default=DEFAULT_LOCALE,
        metavar="country",
        type=str,
        help=f"Country to use for national holidays (default: {DEFAULT_LOCALE})",
    )
    parser.add_argument(
        "--region",
        metavar="region",
        type=str,
        help="State, province or other subdivision within a country",
    )
    parser.add_argument(
        "year",
        default=[DEFAULT_YEAR],
        nargs="*",
        type=valid_year,
        help="years to generate a calendar for (default: current year)",
    )
    return parser


@dataclass
class Calendar(ABC):
    start_month: int = 1
    weekends: list = field(default_factory=lambda: DEFAULT_WEEKENDS)
    no_holiday_weekends: bool = False
    holidays: HolidayBase = None
    filename: str = None

    def __post_init__(self):
        no_border = {x: (0.0, (0, 0, 0), "none") for x in ["top", "right", "bottom", "left"]}
        solid_border = {x: (1.0, (0, 0, 0), "solid") for x in ["top", "right", "bottom", "left"]}
        self.add_style(bg_color=(146, 146, 146), border=solid_border, name="Weekend")
        self.add_style(border=solid_border, name="Month Day")
        self.add_style(bg_color=(0, 0, 0), border=solid_border, name="Holiday")
        self.add_style(
            font_size=10.0,
            bold=True,
            alignment=("center", "middle"),
            border=solid_border,
            name="Year",
        )
        self.add_style(
            font_size=10.0, alignment=("left", "middle"), border=solid_border, name="Month"
        )
        self.add_style(
            font_size=10.0, alignment=("center", "middle"), border=solid_border, name="Day Number"
        )
        self.add_style(border=no_border, name="Empty")
        self.add_style(
            border={"right": (0.0, (0, 0, 0), "none"), "bottom": (0.0, (0, 0, 0), "none")},
            name="Empty Day",
        )

    @abstractmethod
    def add_style(self, **kwargs):
        pass

    @abstractmethod
    def add_sheet(self, sheet_name: str):
        pass

    @abstractmethod
    def col_width(self, sheet: object, col_num: int, width: float):
        pass

    @abstractmethod
    def merge_cells(
        self, sheet: object, row_start: int, col_start: int, row_end: int, col_end: int
    ):
        pass

    @abstractmethod
    def row_height(self, sheet: object, col_num: int, height: float):
        pass

    @abstractmethod
    def save(self):
        pass

    @abstractmethod
    def set_cell_style(self, sheet: object, row_num: int, col_num: int, style: str):
        pass

    @abstractmethod
    def write(self, sheet: object, row_num: int, col_num: int, value: str, style: str):
        pass

    def add_year(self, year: int):
        sheet = self.add_sheet(sheet_name=self.sheet_name(year))
        self.set_cell_sizes(sheet)
        self.set_months(sheet, year)
        self.set_days(sheet, year)

    def set_cell_sizes(self, sheet):
        """Set the row and column sizes for the calendar"""
        for row_num in range(0, 14):
            self.row_height(sheet, row_num, 30.0)
        self.col_width(sheet, 0, 40.0)
        self.col_width(sheet, 1, 60.0)
        self.col_width(sheet, 2, 20.0)
        for col_num in range(3, 34):
            self.col_width(sheet, col_num, 30.0)

    def set_months(self, sheet: object, year: int):
        """Set the borders and merge the month and year names"""

        # Empty cells
        for col_num in range(0, 3):
            self.set_cell_style(sheet, 0, col_num, "Empty")
        for row_num in range(0, 12):
            self.set_cell_style(sheet, row_num + 1, 2, "Empty")

        for offset in range(0, 12):
            self.set_cell_style(sheet, offset + 1, 0, "Month")
            self.set_cell_style(sheet, offset + 1, 1, "Month")

        if self.start_month > 1:
            offset = 13 - self.start_month
            self.merge_cells(sheet, 1, 0, offset, 0)
            self.write(sheet, 1, 0, str(year), "Year")
            self.merge_cells(sheet, offset + 1, 0, 12, 0)
            self.write(sheet, offset + 1, 0, str(year + 1), "Year")
        else:
            self.merge_cells(sheet, 1, 0, 12, 0)
            self.write(sheet, 1, 0, str(year), "Year")

        for month_num in range(0, 12):
            if self.start_month + month_num > 12:
                month_name = calendar.month_name[(self.start_month + month_num) % 12]
            else:
                month_name = calendar.month_name[(self.start_month + month_num)]
            self.write(sheet, month_num + 1, 1, month_name, "Month")

        # Days along top of sheet
        for offset in range(0, 31):
            self.write(sheet, 0, offset + 3, str(offset + 1), "Day Number")

    def set_days(self, sheet: object, year: int):
        """Set the styles and borders for all days of the month"""

        for row_num in range(0, 12):
            month_num = self.start_month + row_num
            if month_num > 12:
                month_dt = date(year, month_num % 12, 1)
                (_, num_days) = calendar.monthrange(year, month_num % 12)
            else:
                month_dt = date(year, month_num, 1)
                (_, num_days) = calendar.monthrange(year, month_num)
            for col_num in range(0, 31):
                row_col = (row_num + 1, col_num + 3)
                if col_num >= num_days:
                    self.set_cell_style(sheet, *row_col, "Empty Day")
                else:
                    day_dt = month_dt + relativedelta(days=col_num)
                    is_weekend = day_dt.isoweekday() in self.weekends
                    is_holiday = self.holidays.get(day_dt) is not None
                    if is_holiday and is_weekend and self.no_holiday_weekends:
                        self.set_cell_style(sheet, *row_col, "Weekend")
                    elif is_holiday:
                        self.set_cell_style(sheet, *row_col, "Holiday")
                    elif is_weekend:
                        self.set_cell_style(sheet, *row_col, "Weekend")
                    else:
                        self.set_cell_style(sheet, *row_col, "Month Day")

    def sheet_name(self, year: int):
        """Year name for sheet, e.g. 2023-24"""
        if self.start_month > 0:
            year_1 = str(year)
            if len(year_1) >= 4:
                year_2 = str(year + 1)[-2:]
            else:
                year_2 = str(year + 1)
            return f"{year_1}-{year_2}"
        else:
            return str(year)


@dataclass
class NumbersCalendar(Calendar):
    def __post_init__(self):
        self.doc = Document(num_header_rows=0, num_header_cols=0, num_rows=13, num_cols=34)
        self.first_sheet = True
        self.styles = {}
        super().__post_init__()

    def add_sheet(self, sheet_name: str) -> object:
        if self.first_sheet:
            self.doc.sheets[0].name = sheet_name
        else:
            self.doc.add_sheet(sheet_name=sheet_name)
        self.first_sheet = False
        return self.doc.sheets[-1]

    def add_style(self, **kwargs):
        self.styles[kwargs["name"]] = kwargs.copy()
        if "border" in kwargs:
            del kwargs["border"]
        self.doc.add_style(**kwargs)

    def col_width(self, sheet: object, col_num: int, width: float):
        sheet.tables[0].col_width(col_num, width)

    def merge_cells(
        self, sheet: object, row_start: int, col_start: int, row_end: int, col_end: int
    ):
        ref = xl_range(row_start, col_start, row_end, col_end)
        sheet.tables[0].merge_cells(ref)

    def row_height(self, sheet: object, row_num: int, height: float):
        sheet.tables[0].row_height(row_num, height)

    def set_cell_style(self, sheet: object, row_num: int, col_num: int, style: str):
        sheet.tables[0].set_cell_style(row_num, col_num, style)
        self.set_border(sheet, row_num, col_num, style)

    def save(self):
        self.doc.save(self.filename)

    def write(self, sheet: object, row_num: int, col_num: int, value: str, style: str):
        sheet.tables[0].write(row_num, col_num, value, style=style)
        self.set_border(sheet, row_num, col_num, style)

    def set_border(self, sheet: object, row_num: int, col_num: int, style: str):
        for side, border in self.styles[style]["border"].items():
            cell = sheet.tables[0].cell(row_num, col_num)
            if not (cell.is_merged and side in ["bottom", "right"]):
                sheet.tables[0].set_cell_border(row_num, col_num, side, Border(*border))


@dataclass
class ExcelCalendar(Calendar):
    def __post_init__(self):
        self.workbook = Workbook(self.filename)
        self.styles = {}
        super().__post_init__()

    def add_sheet(self, sheet_name: str) -> object:
        worksheet = self.workbook.add_worksheet(sheet_name)
        return worksheet

    def add_style(self, **kwargs):
        if "alignment" in kwargs:
            kwargs["align"] = kwargs["alignment"][0]
            kwargs["valign"] = kwargs["alignment"][1]
            del kwargs["alignment"]
        if "bg_color" in kwargs:
            kwargs["bg_color"] = "#{0:02x}{1:02x}{2:02x}".format(
                kwargs["bg_color"][0], kwargs["bg_color"][1], kwargs["bg_color"][2]
            )
        for side, border in kwargs["border"].items():
            if border[2] == "none":
                kwargs[side] = 0
            else:
                kwargs[side] = 1
        del kwargs["border"]
        name = kwargs["name"]
        del kwargs["name"]
        self.styles[name] = self.workbook.add_format(kwargs)

    def col_width(self, sheet: object, col_num: int, width: float):
        sheet.set_column_pixels(col_num, col_num, width)

    def merge_cells(
        self, sheet: object, row_start: int, col_start: int, row_end: int, col_end: int
    ):
        sheet.merge_range(row_start, col_start, row_end, col_end, None, self.styles["Year"])

    def row_height(self, sheet: object, row_num: int, height: float):
        sheet.set_row_pixels(row_num, height)

    def set_cell_style(self, sheet: object, row_num: int, col_num: int, style: str):
        sheet.write(row_num, col_num, None, self.styles[style])

    def save(self):
        self.workbook.close()

    def write(self, sheet: object, row_num: int, col_num: int, value: str, style: str):
        sheet.write(row_num, col_num, value, self.styles[style])


def main():
    alpha2_to_country, country_to_alpha2, alpha2_regions = generate_country_lookups()

    parser = command_line_parser()
    args = parser.parse_args()

    if args.version:
        print(__version__)
    elif args.list_countries:
        for alpha2, name in alpha2_to_country.items():
            print(f"{alpha2}: {name}")
        exit(0)
    elif args.list_regions and args.country is None:
        parser.error("--list-regions requires a country")
    elif args.list_regions:
        if args.country in country_to_alpha2:
            country = country_to_alpha2[args.country]
        elif args.country not in alpha2_to_country:
            parser.error(f"country '{args.country}' not available")
        else:
            country = args.country
            for region in alpha2_regions[country]:
                print(region)
            exit(0)
    else:
        if args.country in country_to_alpha2:
            country = country_to_alpha2[args.country]
        else:
            country = args.country

        if args.output is None:
            filename = "calendar.numbers" if args.format == "numbers" else "calendar.xlsx"
        else:
            filename = args.output

        cls = NumbersCalendar if args.format == "numbers" else ExcelCalendar

        doc = cls(
            start_month=args.start_month,
            weekends=args.weekend,
            no_holiday_weekends=args.no_holiday_weekends,
            holidays=country_holidays(country, subdiv=args.region),
            filename=filename,
        )
        for year in args.year:
            doc.add_year(year)
        doc.save()


if __name__ == "__main__":
    # execute only if run as a script
    main()
