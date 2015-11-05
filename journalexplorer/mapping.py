from recordclass import recordclass

"""
    ColumnMap maps the JournalEntry field to the corresponding Excel column.

    This is the JournalEntry    -> Excel column names    : Column Index
    -------------------------------------------------------------------
    ArticleTitle                -> Title                 :  0
    ArticleAuthors              -> Authors               :  2
    # ArticleCorrespondenceAuthor -> correspondenceAuthor  :  7
    ArticleAbstract             -> Abstract              :  1
    ArticleSubjectTerms         -> subjectTerms          : 10
    JournalTitle                -> pubtitle              : 5
    JournalDate                 -> pubdate               : 4
    # JournalCountry              -> publisher             : 20
    JournalIssue                -> issue                 : 7
    JournalVolume               -> volume                : 8
    JournalYear                 -> year                  : 6
"""
ColumnMap = {
    "ArticleTitle":                 0,
    "ArticleAuthors":               2,
    # "ArticleCorrespondenceAuthor":  7,
    "ArticleAbstract":              1,
    "ArticleSubjectTerms":         10,
    "JournalTitle":                5,
    "JournalDate":                 4,
    # "JournalCountry":              20,
    "JournalIssue":                7,
    "JournalVolume":               8,
    "JournalYear":                 6
}


"""
    The JournalEntry named tuple holds a single publication record.
"""
JournalEntry = recordclass("JournalEntry", "ArticleTitle ArticleAuthors \
                                           ArticleAbstract ArticleSubjectTerms JournalTitle \
                                           JournalDate JournalIssue JournalVolume \
                                           JournalYear")
