# extract-urls
.Synopsis
    Search through exported message files under a specified path, recursively.  Ouput text file reports with unique
    URLs and message metadata.

    Scanning 10,000 may take 10-15 minutes.  A progress bar is displayed.  Selecting the -verboseOutput and
    -writeCSV flags may add another 10 minutes or so.

 .Description
    This script will extract any URLs along with their accompanying text description,
    in addition to the sender, sender IP, recipient, subject (message export filename defaults to subject), 
    and timestamp of message.

    If you use office 365, messages can be exported using the content search in the security and compliance center.
    Make sure to use the `export individual messages` option - this script cannot open .PST files.

    You can also specify a list of URLs to filter out; a filter list with http://www.facebook.com will filter out 
    http://www.facebook.com/user/jim, but not https://www.facebook.com. 

    You must specify to output in either CSV or JSON format.  Note that exporting the detailed report in CSV format 
    is very time consuming.  JSON is much faster.

    Parameters:
    -path:  path to traverse looking for messages.
    -FilterJunkFolders:     flag - include to ignore messages found in a users Junk E-Mail folder.
    -URLFilterList:        full path to urlfilters.conf file - one entry per line, no punctuation.
    -writeCSV:          output in CSV format. '|' used as a delimiter.
    -writeJSON:         output in JSON format.  Faster, epsecially for verbose output.
    -verboseOutput      flag - include to also output a report with the metadata and URLs found in each email scanned.


 .Example
    .\extract-urls.pl1 -path c:\out\export -writeJSON 

    Search through all message files under c:\out\export, and write report of unique URLs in .json format.

 .Example
    PS C:\users\jim>  ~\documents\repos\extract-urls\extract-urls.ps1 -path c:\export\search1 -writeJSON -URLFilterList ~\documents\repos\extract-urls\urlfilters.conf

    Search through all message files under c:\export\search1, write report in .json format, excluding anything that 
    matches an entry in the text file at urlfilters.conf

 .Example
    .\extract-urls.pls1 -path c:\out\export -writeJSON -verboseOutput -URLFilterList c:\users\jim\documents\filters.txt

    Search through all message files in c:\out\export, write a report of uniuqe URLs and metadata, as well as a detailed report listing all messages scanned with URLs.  The specified URL Filter is applied to both reports.
