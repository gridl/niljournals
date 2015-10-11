from optparse import OptionParser
from collections import namedtuple

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
        raise Exception("Wrong numer of arguments")

    # parse keywords
    keywords = options.keyword_filter.replace(' ','').lower().split(",")

    # parse year filter (range)
    years_str = options.year_filter.replace(' ','').split("-")
    if len(years_str) == 1:
        years_int = range(int(years_str[0]))
    elif len(years_str) == 2:
        years_int = range(int(years_str[0]), int(years_str[1]))
    else:
        raise Exception("Wrong year filter format. Either use a single value (2001) or range (2001-2005)")

    return Options(
        JournalFilter=options.journal_filter,
        YearFilter=years_int,
        KeywordFilter=keywords,
        OutputFile=options.output_file,
        InputFile=remainder[0]
    )
