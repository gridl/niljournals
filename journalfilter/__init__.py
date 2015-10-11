
import sys
import xlrd

from optparse import OptionParser
from collections import namedtuple
from mapping import JournalEntry

# Options holds the command line options
Options = namedtuple("Options", "JournalFilter YearFilter KeywordFilter OutputFile InputFile")

# Parses the command line options
def parse_commandline():
    parser = OptionParser(usage='%prog [options] <Excel File>')

    parser.add_option('-j', '--journal-filter',
                      dest="journal_filter",
                      default=None,
                      help="Limit output to a sepcific set of journals"
                      )
    parser.add_option('-y', '--year-filter',
                      dest="year_filter",
                      default=None,
                      help="Limit output to a sepcific year or year range"
                      )
    parser.add_option('-k', '--keyword-filter',
                      dest="keyword_filter",
                      default=None,
                      help="Limit output to a sepcific set of keywords"
                      )
    parser.add_option('-o', '--output',
                      dest="output_file",
                      default=None,
                      help="Write results to file instead of screen"
                      )

    options, remainder = parser.parse_args()

    if len(remainder) != 1:
        parser.print_help()
        sys.exit(-1)

    return Options(
        JournalFilter=options.journal_filter,
        YearFilter=options.year_filter,
        KeywordFilter=options.keyword_filter,
        OutputFile=options.output_file,
        InputFile=remainder[0]
    )


def main():
    """Entry point for the application script"""

    options = parse_commandline()
