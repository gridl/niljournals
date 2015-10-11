from recordclass import recordclass

"""
    ColumnMap maps the JournalEntry field to the corresponding Excel column.

    This is the JournalEntry    -> Excel column names    : Column Index
    -------------------------------------------------------------------
    ArticleTitle                -> Title                 :  0
    ArticleAuthors              -> Authors               :  5
    ArticleCorrespondenceAuthor -> correspondenceAuthor  :  7
    ArticleAbstract             -> Abstract              :  1
    ArticleSubjectTerms         -> subjectTerms          : 31
    JournalTitle                -> pubtitle              : 17
    JournalDate                 -> pubdate               : 16
    JournalCountry              -> publisher             : 20
    JournalIssue                -> issue                 : 12
    JournalVolume               -> volume                : 21
    JournalYear                 -> year                  : 19
"""
ColumnMap = {
    "ArticleTitle":                 0,
    "ArticleAuthors":               5,
    "ArticleCorrespondenceAuthor":  7,
    "ArticleAbstract":              1,
    "ArticleSubjectTerms":         31,
    "JournalTitle":                17,
    "JournalDate":                 16,
    "JournalCountry":              20,
    "JournalIssue":                12,
    "JournalVolume":               21,
    "JournalYear":                 19
}


"""
    The JournalEntry named tuple holds a single publication record.
"""
JournalEntry = recordclass("JournalEntry", "ArticleTitle ArticleAuthors ArticleCorrespondenceAuthor \
                                           ArticleAbstract ArticleSubjectTerms JournalTitle \
                                           JournalDate JournalCountry JournalIssue JournalVolume \
                                           JournalYear")
