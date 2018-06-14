fetch-imap-emails
=================

This is a Python script for fetching email headers from an IMAP mailbox as a CSV file.

It is build using the standard libraries available, but probably requires Python 3.6 to function.


Example
-------

```
fetch-imap-emails.py --fields=date:isotime,from:address,subject "imaps://user:$PASSWORD@imap.example.com/INBOX" >messages.csv
```
