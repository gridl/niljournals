
import sys
import xlrd
import pprint

from mapping import JournalEntry, ColumnMap as CM
from cmdlineparser import Options, parse_commandline


def main():
    """Entry point for the application script"""

    # Parse the command line
    try:
        options = parse_commandline()

    except Exception, ex:
        print str(ex)
        sys.exit(-1)

    # Open the Excel file via xlrd
    try:
        book = xlrd.open_workbook(options.InputFile)
        # for now we always use the first sheet of the file
        sh = book.sheet_by_index(0)
        print " * Using Excel sheet '%s' with %d rows as input data." % (sh.name, sh.nrows)

    except Exception, ex:
        print str(ex)
        sys.exit(-1)

    # Print options again so the user can double-check
    if options.OutputFile == None:
        output = "screen"
    else:
        output = "file '%s'" % options.OutputFile

    print " * Generating output dataset with the following options:"
    print "   * Journal filter  : %s" % (options.JournalFilter)
    print "   * Keyword filter  : %s" % (options.KeywordFilter)
    print "   * Year filter     : %s" % (options.YearFilter)
    print "   * Output to       : %s" % (output)

    # Now we can do the actual parsing

    for rx in range(1,sh.nrows):
        # print "%s : %s" % (rx, sh.row(0)[rx])
        row = sh.row(rx)

        je = JournalEntry(
            ArticleTitle=row[CM['ArticleTitle']],
            ArticleAuthors=row[CM['ArticleAuthors']],
            ArticleCorrespondenceAuthor=row[CM['ArticleCorrespondenceAuthor']],
            ArticleAbstract=row[CM['ArticleAbstract']],
            ArticleSubjectTerms=row[CM['ArticleSubjectTerms']],
            JournalTitle=row[CM['JournalTitle']],
            JournalDate=row[CM['JournalDate']],
            JournalCountry=row[CM['JournalCountry']],
            JournalIssue=row[CM['JournalIssue']],
            JournalVolume=row[CM['JournalVolume']],
            JournalYear=row[CM['JournalYear']])
        pp = pprint.PrettyPrinter(depth=6)
        pp.pprint(je)
