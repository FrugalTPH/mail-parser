# mail-parser

A console-based service I built to parse emails of msg & eml formats. The tool extracts various pieces of meta-data as well as the email body & writes them into a json format version of the email, which is archived alongside the source file. 

The json version of the email the provides a few uses:

1. Provides the email's meta-data for use in external databases.
2. Allows the email to become indexable/searchable on dropbox.

This service is used in my other project "FMMS".
