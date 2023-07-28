import argparse
import dateutil

from calendar import monthrange
from locale import getlocale
from numbers_calendar import __version__
from numbers_parser import Document, Border, RGB, Style
from datetime import date
from dateutil import parser

DEFAULT_FILENAME = "calendar.numbers"
DEFAULT_LOCALE = getlocale()[0][-2:]


def valid_date(s):
    try:
        dt = dateutil.parse(s)
    except dateutil.ParseError as e:
        raise argparse.ArgumentTypeError(str(e))


def command_line_parser():
    parser = argparse.ArgumentParser(
        description="Create Apple Numbers spreadsheet calendars using python"
    )
    parser.add_argument("-V", "--version", action="store_true")
    parser.add_argument(
        "--start-date",
        type=valid_date,
        default=date(date.today().year, 1, 1),
        help="Calendar start date (default: Jan 1 of current year)",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=DEFAULT_FILENAME,
        help=f"output file (default: {DEFAULT_FILENAME})",
    )
    parser.add_argument(
        "--country",
        default=DEFAULT_LOCALE,
        help=f"Country to use for national holidays (default: {DEFAULT_LOCALE})",
    )
    return parser


def main():
    parser = command_line_parser()
    args = parser.parse_args()

    if args.version:
        print(__version__)
    else:
        doc = Document()
        table = doc.sheets[0].tables[0]
        for month_num in range(0, 12):
            num_days = monthrange(args.start_date.year, args.start_date.month)[1]
            for col_num in range(0, 31):
                if col_num >= num_days:
                    continue
                dt = date(args.start_date.year, month_num + 1, col_num + 1)
                border = Border(1.0, RGB(0, 0, 0), "solid")
                table.add_border(month_num, col_num, ["top", "right", "bottom", "left"], border)
                if dt.isoweekday() == 6 or dt.isoweekday() == 7:
                    table.set_cell_style(month_num, col_num, Style(bg_color=RGB(0, 0, 0)))
        doc.save(args.output)


if __name__ == "__main__":
    # execute only if run as a script
    main()
