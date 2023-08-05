# spreadsheet-calendar

A command-line utility to generate whole year calendars in Excel and Apple Numbers formats. Each year is saved as a separate sheet and is formatted to include national holidays.

## Installation

`spreadsheet-calendar` is most simply installed using pip:

``` text
python3 -m pip install spreadsheet-calendar
```

## Usage

By default, `spreadsheet-calendar` creates a calendar for the current year and the current locale. Available national holidays are those supported by [python-holidays](https://pypi.org/project/holidays/). The list of available countres can be listed using `--list-countries`. Within a country, subdivisions such as regions and states can be used. These can be listed using `--country=name --list-regions`. When listing countries and regions, no calendar is generated and `spreadsheet-calendar` exits after printing the `stdout`.

``` text
spreadsheet-calendar [-h] [-V] [--list-countries] [--list-regions]
                     [--no-holiday-weekends] [--format {numbers,excel}]
                     [--start-month month] [--weekend [day ...]] [-o filename]
                     [--country country] [--region region]
                     [year ...]
```

Multiple years can be passed as a positional argument. The default year is the current year.

Available options:

* `--help`: print the command-line usage and exit.
* `-V`, `--version`: print the version of `spreadsheet-calendar` and exit.
* `--list-countries`: list the supported countries, mapping 2-character locale code to country name.
* `--list-regions`: list the supported regions within a country. If `--country` is not given, the current locale of the script is used.
* `--no-holiday-weekends`: if present, weekends are always filled grey even when the day is a holiday. Default is that holidays are black filled regardless of weekdays or weekends.
* `--format`: the output format for the calendar. Choices are `excel` for Microsoft Excel's XLSX format (the default) or `numbers` for Apple Numbers format.
* `--start-month`: the month for the first row of the calendar (default: January). Month names can be short names like `Jan` or full names like `January`.
* `--weekend`: days of the week to be marked as weekends (default: Saturday and Sunday). Multiple days can be passed on the command line, for example `--weekend=Fri --weekend=Sat`. Day names can be short names like `Fri` or full names like `Friday`.
* `-o`, `--output`: the nme of the output file. Files are always overwitten. The default depends upon the `--format` option: `calendar.xlsx` for `--format=excel` and `calendar.numbers` for `--format=numbers`.
* `--country`: the name of the country to use for national holidays (default: current locale's). Country names can be 2-character locale names like `GB` or full country names like `United Kingdom`.
* `--region`: the name of the subdivision within a country. Names can be the abbreviations supportef by [python-holidays](https://pypi.org/project/holidays/) like `ENG` or full names like `England`.

## License

All code in this repository is licensed under the [MIT License](https://github.com/masaccio/numbers-calendar/blob/master/LICENSE)
