
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
        print "\n * Using Excel sheet '%s' with %d rows of data." % (sh.name, sh.nrows)

    except Exception, ex:
        print str(ex)
        sys.exit(-1)

    # Print options again so the user can double-check
    if options.OutputFile == None:
        output = "screen"
    else:
        output = "file '%s'" % options.OutputFile

    print "\n * Generating output dataset with the following options:"
    print "   * Journal filter  : %s" % (options.JournalFilter)
    print "   * Keyword filter  : %s" % (options.KeywordFilter)
    print "   * Year filter     : %s" % (options.YearFilter)
    print "   * Output to       : %s" % (output)

    # Now we can do the actual parsing. First we sanitize and throw all the data into a record
    # list and do some
    journals = list()

    for rx in range(1,sh.nrows):
        row = sh.row(rx)

        je = JournalEntry(
            ArticleTitle=row[CM['ArticleTitle']].value,
            ArticleAuthors=row[CM['ArticleAuthors']].value,
            ArticleCorrespondenceAuthor=row[CM['ArticleCorrespondenceAuthor']].value,
            ArticleAbstract=row[CM['ArticleAbstract']].value,
            ArticleSubjectTerms=row[CM['ArticleSubjectTerms']].value,
            JournalTitle=row[CM['JournalTitle']].value,
            JournalDate=row[CM['JournalDate']].value,
            JournalCountry=row[CM['JournalCountry']].value,
            JournalIssue=row[CM['JournalIssue']].value,
            JournalVolume=row[CM['JournalVolume']].value,
            JournalYear=row[CM['JournalYear']].value)

        # Remove the '*' from ArticleSubjectTerms
        je.ArticleSubjectTerms = je.ArticleSubjectTerms.replace("*", "")
        je.JournalYear = int(je.JournalYear)

        journals.append(je)

    # print journals
