import argparse
import calendar
from datetime import date
from locale import getlocale

from dateutil.relativedelta import relativedelta
from numbers_parser import RGB, Border, Document, Style, Alignment

from numbers_calendar import __version__


def generate_month_map():
    month_map = {}
    for mon in range(1, 13):
        month_map[calendar.month_name[mon]] = mon
        month_map[calendar.month_abbr[mon]] = mon
    return month_map


DEFAULT_FILENAME = "calendar.numbers"
DEFAULT_LOCALE = getlocale()[0][-2:]
MONTH_MAP = generate_month_map()
ALL_BORDERS = ["top", "right", "bottom", "left"]
SOLID_BORDER = Border(1.0, RGB(0, 0, 0), "solid")
NO_BORDER = Border(0.0, RGB(0, 0, 0), "none")


def valid_month(month):
    try:
        _ = MONTH_MAP[month]
    except KeyError:
        raise argparse.ArgumentTypeError(f"'{month}' is not a valid month abbreviation or name")
    return MONTH_MAP[month]


def valid_year(year):
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
        "--start-month",
        type=valid_month,
        default="Jan",
        help="Start month for calendar (default: Jan)",
    )
    parser.add_argument(
        "-o",
        "--output",
        metavar="filename",
        default=DEFAULT_FILENAME,
        help=f"output file (default: {DEFAULT_FILENAME})",
    )
    parser.add_argument(
        "--country",
        default=DEFAULT_LOCALE,
        metavar="country",
        help=f"Country to use for national holidays (default: {DEFAULT_LOCALE})",
    )
    parser.add_argument("year", nargs="+", type=valid_year, help="years to generate a calendar for")
    return parser


def sheet_name(year, start_month):
    if start_month > 0:
        year_1 = str(year)
        if len(year_1) >= 4:
            year_2 = str(year + 1)[-2:]
        else:
            year_2 = str(year + 1)
        return f"{year_1}-{year_2}"
    else:
        return str(year)


def create_calendar(args):
    doc = Document()
    doc.sheets[0].name = sheet_name(args.year[0], args.start_month)
    for year in args.year[1:]:
        doc.add_sheet(name=sheet_name(year, args.start_month))

    align_cm = Alignment("center", "middle")
    align_lm = Alignment("left", "middle")
    weekend_style = doc.add_style(bg_color=RGB(0, 0, 0), name="Weekend")
    year_style = doc.add_style(bold=True, font_size=10.0, alignment=align_cm, name="Year")
    month_style = doc.add_style(font_size=10.0, alignment=align_lm, name="Month")
    day_num_style = doc.add_style(font_size=10.0, alignment=align_cm, name="Day Number")

    for year in args.year:
        table = doc.sheets[sheet_name(year, args.start_month)].tables[0]

        for row_num in range(0, 14):
            table.row_height(row_num, 30.0)
        table.col_width(0, 40.0)
        table.col_width(1, 60.0)
        table.col_width(2, 20.0)
        for col_num in range(3, 34):
            table.col_width(col_num, 30.0)

        # Month names cells
        table.set_cell_border(0, 0, ALL_BORDERS, NO_BORDER)
        table.set_cell_border(0, 1, ALL_BORDERS, NO_BORDER)
        for row_num in range(0, 14):
            table.set_cell_border(row_num, 2, ALL_BORDERS, NO_BORDER)
        table.num_header_rows = 0
        table.num_header_cols = 0
        if args.start_month > 1:
            year_1_length = 13 - args.start_month
            table.merge_cells(f"A2:A{year_1_length + 1}")
            table.write("A2", str(year), style=year_style)
            table.merge_cells(f"A{year_1_length + 2}:A13")
            table.write(f"A{year_1_length + 1}", str(year + 1), style=year_style)
            table.set_cell_border("A2", ALL_BORDERS, SOLID_BORDER)
            table.set_cell_border(f"A{year_1_length + 2}", ALL_BORDERS, SOLID_BORDER)
        else:
            table.merge_cells("A2:A13")
            table.write("A2", str(year), style=year_style)
            table.set_cell_border("A2", ALL_BORDERS, SOLID_BORDER)

        for month_num in range(0, 12):
            if args.start_month + month_num > 12:
                month_name = calendar.month_name[(args.start_month + month_num) % 12]
            else:
                month_name = calendar.month_name[(args.start_month + month_num)]
            table.write(month_num + 1, 1, month_name, style=month_style)
            table.set_cell_border(month_num + 1, 1, ALL_BORDERS, SOLID_BORDER)

        for offset in range(0, 31):
            table.write(0, offset + 3, str(offset + 1), style=day_num_style)
            table.set_cell_border(0, offset + 3, ALL_BORDERS, SOLID_BORDER)

        start_dt = date(year, args.start_month, 1)
        for row_num in range(0, 12):
            month_num = args.start_month + row_num
            if month_num > 12:
                month_dt = start_dt + relativedelta(months=month_num % 12, years=1)
                (_, num_days) = calendar.monthrange(year, month_num % 12)
            else:
                month_dt = start_dt + relativedelta(months=month_num)
                (_, num_days) = calendar.monthrange(year, month_num)
            for col_num in range(0, 31):
                if col_num >= num_days:
                    table.set_cell_border(row_num + 1, col_num + 3, "right", NO_BORDER)
                else:
                    day_dt = month_dt + relativedelta(days=col_num)
                    table.set_cell_border(row_num + 1, col_num + 3, ALL_BORDERS, SOLID_BORDER)
                    # if day_dt.isoweekday() == 6 or day_dt.isoweekday() == 7:
                    #     table.set_cell_style(month_num, col_num, weekend_style)

    doc.save(args.output)


def main():
    parser = command_line_parser()
    args = parser.parse_args()

    if args.version:
        print(__version__)
    else:
        create_calendar(args)


if __name__ == "__main__":
    # execute only if run as a script
    main()
