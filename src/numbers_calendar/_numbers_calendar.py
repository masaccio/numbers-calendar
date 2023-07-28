import argparse
import dateutil
import sys

from numbers_calendar import __version__
from numbers_parser import Document
from datetime import date
from dateutil import parser


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
        "--start",
        type=valid_date,
        default=date(date.today().year, 1, 1),
        help="Calendar start date (default: Jan 1 of current year)",
    )
    parser.add_argument(
        "-o",
        "--output",
        # default="calendar.numbers",
        help="output file (default: calendar.numbers)",
    )
    return parser


def main():
    parser = command_line_parser()
    args = parser.parse_args()

    if args.version:
        print(__version__)
    elif len(args.document) == 0:
        parser.print_help()
    else:
        pass


if __name__ == "__main__":
    # execute only if run as a script
    main()
