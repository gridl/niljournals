
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

    print "\n * Generating keyword timeline with the following options:"
    print "   * Abstract filter : %s" % (options.AbstractFilter)
    print "   * Keyword filter  : %s" % (options.KeywordFilter)
    print "   * Year filter     : %s" % (options.YearFilter)

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
        je.ArticleSubjectTerms = je.ArticleSubjectTerms.replace("*", "").replace(" ", "").lower().split(",")
        je.JournalYear = int(je.JournalYear)

        je.ArticleTitle = je.ArticleTitle.lower()
        je.ArticleAbstract = je.ArticleAbstract.lower()

        journals.append(je)

    output = dict()

    for journal in journals:

        # match year

        if options.YearFilter is not None:
            if journal.JournalYear not in options.YearFilter:
                break

        # match keywords
        keyword_match = True;
        if options.KeywordFilter is not None:
            for kw in options.KeywordFilter:
                if kw not in journal.ArticleSubjectTerms:
                    keyword_match = False;

        # match abstracts
        if options.AbstractFilter is not None:
            abstract_match = True;
            for akw in options.AbstractFilter:
                if akw not in journal.ArticleAbstract:
                    abstract_match = False;

        if (options.AbstractFilter is not None) and (abstract_match == True):
            if journal.JournalYear in output:
                output[journal.JournalYear]['counter'] += 1
                output[journal.JournalYear]['articles'].append(journal.ArticleTitle)
            else:
                output[journal.JournalYear] = dict()
                output[journal.JournalYear]['counter'] = 1
                output[journal.JournalYear]['articles'] = list()
                output[journal.JournalYear]['articles'].append(journal.ArticleTitle)


            # output.append(journal)
            # pprint.pprint( journal )
            # print ""



    pprint.pprint(output)

    print "\n Results: %d" % len(output)
