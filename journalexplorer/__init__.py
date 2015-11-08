
import sys
import xlrd
import json
import pprint

from mapping import JournalEntry, ColumnMap as CM
from cmdlineparser import Options, parse_commandline

##########################################################################################
#
#
def main():

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

    # parse abstract filters
    if options.AbstractFilter is not None:
        abstracts = options.AbstractFilter.lower().split(",")

    else:
        abstracts = None

    print "\n * Generating keyword timeline with the following options:"
    print "   * Abstract filter : %s" % (abstracts)
    print "   * Keyword filter  : %s" % (options.KeywordFilter)
    print "   * Year filter     : %s" % (options.YearFilter)

    output = analyze(options.InputFile, options.ShowTitles, abstracts, options.KeywordFilter, options.YearFilter)

    formatted_json = json.dumps(output, sort_keys=False, indent=4)

    from pygments import highlight, lexers, formatters
    colorful_json = highlight(unicode(formatted_json, 'UTF-8'), lexers.JsonLexer(), formatters.TerminalFormatter())
    print(colorful_json)


##########################################################################################
#
#
def analyze(inputfile, show_titles, abstract_filter, keyword_filter, year_filter):
    """Entry point for the application script"""

    # Open the Excel file via xlrd
    try:
        book = xlrd.open_workbook(inputfile)
        # for now we always use the first sheet of the file
        sh = book.sheet_by_index(0)
        print "\n * Using Excel sheet '%s' with %d rows of data." % (sh.name, sh.nrows)

    except Exception, ex:
        print str(ex)
        sys.exit(-1)

    # Now we can do the actual parsing. First we sanitize and throw all the data into a record
    # list and do some
    articles = list()

    for rx in range(1,sh.nrows):
        row = sh.row(rx)

        # skip empty rows
        if row[CM['ArticleTitle']].value == '':
            continue

        je = JournalEntry(
            ArticleTitle=row[CM['ArticleTitle']].value,
            ArticleAuthors=row[CM['ArticleAuthors']].value,
            # ArticleCorrespondenceAuthor=row[CM['ArticleCorrespondenceAuthor']].value,
            ArticleAbstract=row[CM['ArticleAbstract']].value,
            ArticleSubjectTerms=row[CM['ArticleSubjectTerms']].value,
            JournalTitle=row[CM['JournalTitle']].value,
            JournalDate=row[CM['JournalDate']].value,
            # articleCountry=row[CM['articleCountry']].value,
            JournalIssue=row[CM['JournalIssue']].value,
            JournalVolume=row[CM['JournalVolume']].value,
            JournalYear=row[CM['JournalYear']].value)

        # Remove the '*' from ArticleSubjectTerms
        je.ArticleSubjectTerms = je.ArticleSubjectTerms.replace("*", "").replace(" ", "").lower().split(",")
        je.JournalYear = int(je.JournalYear)

        je.ArticleTitle = je.ArticleTitle.lower()
        je.ArticleAbstract = je.ArticleAbstract.lower()

        articles.append(je)

    output = dict()

    # loop over all articles
    for article in articles:

        jname = article.JournalTitle
        year = article.JournalYear

        # Create year if it doesn't exist
        if year not in output:
            output[year] = dict()

        # Creat journal if it doesn't exits
        if article.JournalTitle not in output[year]:
            output[year][jname] = dict()
            output[year][jname]['total'] = 0
            output[year][jname]['matches'] = 0

            if show_titles:
                output[year][jname]['titles'] = list()

        # Count the total articles per journal per year
        output[year][jname]['total'] += 1

        # Match abstracts
        if abstract_filter is not None:
            abstract_match = True;
            for akw in abstract_filter:
                if akw not in article.ArticleAbstract:
                    abstract_match = False;

        if (abstract_filter is not None) and (abstract_match == True):
            output[year][jname]['matches'] += 1
            if show_titles:
                output[year][jname]['titles'].append(article.ArticleTitle)

        # Calculate perentages
        for year in output:
            for journal in output[year]:
                entry = output[year][journal]
                entry['matches_percent'] = 100.0 / entry['total'] * entry['matches']

    return output
