import argparse
import calendar
from datetime import date
from locale import getlocale

from dateutil.relativedelta import relativedelta
from holidays import country_holidays
from numbers_parser import RGB, Border, Document, Alignment
from os.path import basename
from sys import argv, exit, stderr

from numbers_calendar import __version__


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


DEFAULT_FILENAME = "calendar.numbers"
DEFAULT_LOCALE = getlocale()[0][-2:]
MONTH_MAP = generate_month_map()
WEEKDAY_MAP = generate_weekday_map()
ALL_BORDERS = ["top", "right", "bottom", "left"]
SOLID_BORDER = Border(1.0, RGB(0, 0, 0), "solid")
NO_BORDER = Border(0.0, RGB(0, 0, 0), "none")


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


def command_line_parser():
    parser = argparse.ArgumentParser(
        description="Create Apple Numbers spreadsheet calendars using python"
    )
    parser.add_argument("-V", "--version", action="store_true")
    parser.add_argument(
        "--no-holiday-weekends",
        action="store_true",
        default=False,
        help="Don't color weekends with holidays",
    )
    parser.add_argument(
        "--start-month",
        type=valid_month,
        default="Jan",
        help="Start month for calendar (default: Jan)",
    )
    parser.add_argument(
        "--weekend",
        type=valid_weekday,
        default=[6, 7],
        nargs="*",
        action="extend",
        metavar="day",
        help="Days to highlight as weekends (default: Sat, Sun)",
    )
    parser.add_argument(
        "-o",
        "--output",
        metavar="filename",
        default=DEFAULT_FILENAME,
        help=f"Output file (default: {DEFAULT_FILENAME})",
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
    parser.add_argument("year", nargs="+", type=valid_year, help="years to generate a calendar for")
    return parser


def sheet_name(year, start_month):
    """Year name for sheet, e.g. 2023-24"""
    if start_month > 0:
        year_1 = str(year)
        if len(year_1) >= 4:
            year_2 = str(year + 1)[-2:]
        else:
            year_2 = str(year + 1)
        return f"{year_1}-{year_2}"
    else:
        return str(year)


def set_calendar_cell_sizes(table):
    """Set the row and column sizes for the calendar"""
    for row_num in range(0, 14):
        table.row_height(row_num, 30.0)
    table.col_width(0, 40.0)
    table.col_width(1, 60.0)
    table.col_width(2, 20.0)
    for col_num in range(3, 34):
        table.col_width(col_num, 30.0)


def set_month_borders(doc, table, year, start_month):
    """Set the borders and merge the month and year names"""
    table.set_cell_border(0, 0, ALL_BORDERS, NO_BORDER)
    table.set_cell_border(0, 1, ALL_BORDERS, NO_BORDER)
    for row_num in range(0, 14):
        table.set_cell_border(row_num, 2, ALL_BORDERS, NO_BORDER)
    table.num_header_rows = 0
    table.num_header_cols = 0
    if start_month > 1:
        year_1_length = 13 - start_month
        year_1_end_ref = f"A{year_1_length + 1}"
        year_2_start_ref = f"A{year_1_length + 2}"
        table.merge_cells(f"A2:{year_1_end_ref}")
        table.write("A2", str(year), style="Year")
        table.merge_cells(f"{year_2_start_ref}:A13")
        table.write(year_1_end_ref, str(year + 1), style="Year")
        table.set_cell_border("A2", ALL_BORDERS, SOLID_BORDER)
        table.set_cell_border(year_2_start_ref, ALL_BORDERS, SOLID_BORDER)
    else:
        table.merge_cells("A2:A13")
        table.write("A2", str(year), style="Year")
        table.set_cell_border("A2", ALL_BORDERS, SOLID_BORDER)


def set_month_names(doc, table, year, start_month):
    """Set the names of months and years"""
    for month_num in range(0, 12):
        if start_month + month_num > 12:
            month_name = calendar.month_name[(start_month + month_num) % 12]
        else:
            month_name = calendar.month_name[(start_month + month_num)]
        table.write(month_num + 1, 1, month_name, style="Month")
        table.set_cell_border(month_num + 1, 1, ALL_BORDERS, SOLID_BORDER)

    for offset in range(0, 31):
        table.write(0, offset + 3, str(offset + 1), style="Day Number")
        table.set_cell_border(0, offset + 3, ALL_BORDERS, SOLID_BORDER)


def set_day_borders(doc, table, year, start_month, weekends, holidays, no_holiday_weekends):
    """Set the borders for all days of the month"""

    for row_num in range(0, 12):
        month_num = start_month + row_num
        if month_num > 12:
            month_dt = date(year, month_num % 12, 1)
            (_, num_days) = calendar.monthrange(year, month_num % 12)
        else:
            month_dt = date(year, month_num, 1)
            (_, num_days) = calendar.monthrange(year, month_num)
        for col_num in range(0, 31):
            if col_num >= num_days:
                table.set_cell_border(row_num + 1, col_num + 3, "right", NO_BORDER)
            else:
                day_dt = month_dt + relativedelta(days=col_num)
                table.set_cell_border(row_num + 1, col_num + 3, ALL_BORDERS, SOLID_BORDER)
                is_weekend = day_dt.isoweekday() in weekends
                is_holiday = holidays.get(day_dt) is not None
                if is_holiday and is_weekend and no_holiday_weekends:
                    table.set_cell_style(row_num + 1, col_num + 3, "Weekend")
                elif is_holiday:
                    table.set_cell_style(row_num + 1, col_num + 3, "Holiday")
                elif is_weekend:
                    table.set_cell_style(row_num + 1, col_num + 3, "Weekend")


def create_calendar(parser, args):
    try:
        holidays = country_holidays(args.country, subdiv=args.region)
    except NotImplementedError as e:
        parser.print_usage()
        script_name = basename(argv[0])
        print(f"{script_name}: error: {e}", file=stderr)
        exit(1)

    doc = Document()
    doc.sheets[0].name = sheet_name(args.year[0], args.start_month)
    for year in args.year[1:]:
        doc.add_sheet(sheet_name=sheet_name(year, args.start_month))

    align_cm = Alignment("center", "middle")
    align_lm = Alignment("left", "middle")
    doc.add_style(bg_color=RGB(146, 146, 146), name="Weekend")
    doc.add_style(bg_color=RGB(0, 0, 0), name="Holiday")
    doc.add_style(font_size=10.0, bold=True, alignment=align_cm, name="Year")
    doc.add_style(font_size=10.0, alignment=align_lm, name="Month")
    doc.add_style(font_size=10.0, alignment=align_cm, name="Day Number")

    for year in args.year:
        table = doc.sheets[sheet_name(year, args.start_month)].tables[0]
        set_calendar_cell_sizes(table)
        set_month_borders(doc, table, year, args.start_month)
        set_month_names(doc, table, year, args.start_month)
        set_day_borders(
            doc, table, year, args.start_month, args.weekend, holidays, args.no_holiday_weekends
        )

    doc.save(args.output)


def main():
    parser = command_line_parser()
    args = parser.parse_args()

    if args.version:
        print(__version__)
    else:
        create_calendar(parser, args)


if __name__ == "__main__":
    # execute only if run as a script
    main()
