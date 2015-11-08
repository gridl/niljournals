from optparse import OptionParser
from collections import namedtuple

# Options holds the command line options
Options = namedtuple("Options", "YearFilter KeywordFilter ShowTitles AbstractFilter InputFile")

# Parses the command line options
def parse_commandline():
    parser = OptionParser(usage='%prog [options] <Excel File>')

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

    parser.add_option('-a', '--abstract-filter',
                      dest="abstract_filter",
                      default=None,
                      help="Limit output to a sepcific set of keywords"
                      )

    parser.add_option('-t', '--show-titles',
                      dest="show_titles",
                      default=None,
                      action="store_true",
                      help="Show titles in addition to the occurence count"
                      )

    options, remainder = parser.parse_args()

    if len(remainder) != 1:
        parser.print_help()
        raise Exception("Wrong numer of arguments")

    # parse keyword filter
    if options.keyword_filter is not None:
        keywords = options.keyword_filter.replace(' ','').lower().split(",")
    else:
        keywords = None

    # parse abstract filters
    if options.abstract_filter is not None:
        #abstracts = options.abstract_filter.replace(' ','').lower().split(",")
        abstracts = options.abstract_filter.lower().split(",")

    else:
        abstracts = None

    # parse year filter (range)
    if options.year_filter is not None:
        years_str = options.year_filter.replace(' ','').split("-")
        if len(years_str) == 1:
            years_int = range(int(years_str[0]))
        elif len(years_str) == 2:
            years_int = range(int(years_str[0]), int(years_str[1]))
        else:
            raise Exception("Wrong year filter format. Either use a single value (2001) or range (2001-2005)")
    else:
        years_int = None

    return Options(
        YearFilter=years_int,
        KeywordFilter=keywords,
        AbstractFilter=abstracts,
        ShowTitles=options.show_titles,
        InputFile=remainder[0]
    )
