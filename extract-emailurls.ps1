 <#
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
  #>
param(
    [switch] $FilterJunkFolders,
    [switch] $writeCSV,
    [switch] $writeJSON,
    [switch] $verboseOutput,
    [Parameter(Mandatory=$true)][String] $path,
    [string] $URLFilterList ="urlfilters.conf")
    
function main($FilterJunkFolders, $path, $urlfilters, $writeCSV, $writeJSON, $verboseOutput) {
    <#
     * walk path and uncompress and .zip files
     * open any .msg files, extract URLS and email metadata
     * export results as csv

     :param: no-junk  {boolean} Ignore emails that were found in Junk Email folders
     :param: path     {string}  Directory containing emails or .zips of emails
     :param: unique   {boolean} Only output unique URLs
     #>

    "`n`n`n`n`n" #make sure progress bar doesn't hide text output
 
    $starttime = get-date
    "Execution begain at: " +$starttime

    #$FilterJunkFolders = $False
    $metadata = New-Object System.Collections.ArrayList
    $errors = New-Object System.Collections.ArrayList
    $msgFiles = New-Object System.Collections.ArrayList
    $uniqueURLs = New-Object System.Collections.ArrayList

    $url_pattern = @'
href=\"(?<url>[a-zA-Z]{3,5}:\/\/[^\"]*)\">(?<text>[^(?=<\/a]*)
'@

    #"Scanning " + $path
    $files = Get-ChildItem -path $path -file -recurse

    #$zips = @()
    #foreach ($file in $files)    {
    #    if ($file.Extension -eq ".zip") {
    #        $zips += $file }
    #}

    foreach ($file in $files) {
        if ($file.Extension -eq ".msg") {
            $msgFiles.add($file) > $null
        }
    }

     
    # TODO uncompress zip files

    foreach ($file in $msgfiles) {
        
        $completion = [math]::Round($msgfiles.indexOf($file)/$files.Length * 100)
        $str = "Parsing " + [String]$msgfiles.Count +" Message Files in " + $path
        Write-Progress -Activity $str -Status "$completion% Complete" -PercentComplete $completion        
          
        try {
            $data = [io.file]::ReadAllText($file.FullName)
            $data = $data -replace '\x00+' 
        }
        catch {
            $errors.add($file.FullName) > Null
            continue
        }

        if ($data -match $url_pattern) {
            $temp2 = Get-Metadata -data $data -url_pattern $url_pattern -urlfilters $urlfilters
            $metadata.add($temp2) > $null #append results but suppress output 
        }
        
    }

    ## reformat data as necessary and write output files

    if ($metadata.Count -gt 0) {
        # Re-format results into a unique list of URLs (with sample metadata)
        Format-UniqueMetadata -metadata $metadata -uniqueURLs $uniqueURLs  

        # write unique output 
        $headers = "Host | URL | Text | Subject| Sender | Date | Recipient | Sender_IP  | Similar_Count"
        
        if ($writeCSV) {
            $csv = $null
            "Format unique.csv"
            Measure-Command -Expression{   
                foreach ($i in $uniqueURLs | Sort-Object) {
                    $csv += $i[0],$i[1][0],$i[1][1] -join " | "
                    $csv += '| '
                    $csv += $i[1][2].values -join ' | '
                    $csv += ' | ' + $i[1][3] 
                    $csv += "`r`n"
                }
            } 
               
            "Write unique.csv"
            measure-command -Expression {
                $str = $path + '\' + "unique.csv"
                Set-Content -path $str -Value $headers
                Add-Content -path $str -Value $csv 
                "Wrote " + $path + "\unique.csv with " + $uniqueURLs.Count + " unique URLs"
            }
        }

        if ($writeJSON) {

            "Format unique.json"
            $json = $null
            measure-command -expression {
                $json = foreach ($i in $uniqueURLs | Sort-Object) {
                    ConvertTo-Json -InputObject $i -Depth 5 | ForEach-Object { [System.Text.RegularExpressions.Regex]::Unescape($_) }
                }  
            }   

            "Write unique.json"
            measure-command -Expression {
                $str = $path + '\' + "unique.json"
                Set-Content -Path $str -Value $json 
                
                "Wrote " + $path +"\unique.json with " + $uniqueURLs.Count + " unique URLs"
            }
        }
    }

    # write detailed output

    if ($writeJSON -and $verboseOutput) {
        "Format data.json"
        measure-command {
            $json = foreach($i in $metadata | Sort-Object) {
                ConvertTo-Json -InputObject $i -Depth 5 | ForEach-Object { [System.Text.RegularExpressions.Regex]::Unescape($_) }
            }
        }

        "Write data.json"
        measure-command {
            $str = $path + '\' + "data.json"
            Set-Content -Path $str -value $json
            "Wrote $str "
        }
    }

    if ($writeCSV -and $verboseOutput) {
    "Format data.csv began at: " + [string](get-date)
        measure-command {
            $str = $path + '\' + "data.csv"   
            $header = "URL| URL_Text| Subject| Sender| Date | Recipient | Sender_IP"
            $csv = $null

            foreach ($i in $metadata | Sort-Object) {

                foreach ($j in $i['links']) {
                    $csv += $j['url'] + ' | ' + $j['text'] + ' | '
                }
                $csv += $i.metadata.values -join ' | '       

                $csv += "`r`n"
            }
        }


        "Write data.csv.  Formatting ended at: " + [string](get-date)
        measure-command {
            set-content -path $str -Value $header
            Add-Content -Path $str -Value $csv
            
            "Wrote $str"
        }
    }
    
    if ($errors.Count -gt 0) {
        $stream = [System.IO.StreamWriter]::new($path + "\errors.csv")
        $stream.writeline("Errors")
        
        $errors| Sort-Object |ForEach-Object {
            
            $stream.WriteLine( $_)}
        "Wrote " + $path + "\errors.csv"
        $stream.Close()
    }
 
    [String]$files.Length + " files found"
    [String]$msgFiles.Count + " message files found"
    [String]$metadata.Count + " messages parsed successfully."
    [String]($msgFiles.Count - $metadata.Count) + " messages had no URL"
    [String]$errors.Count + " messages could not be opened."
    $endtime = get-date
    "Execution ended at: " + $endtime
    "Execution Duration: {0:HH:mm:ss}" -f ([datetime]($endtime - $starttime).Ticks) 
 }

function Format-DateTime($string) {
    try {
        $a = ([DateTime]$string).tostring("u") }
    catch 
    {"format-datetime failed for: " + $string}
    return $a
}


function Format-UniqueMetadata($metadata, $uniqueURLs) {
    <#
    * reformat $metadata into list of unique urls instead of one result per message

    :param: $metadta
    :param: $uniqueURLs

    #>

    foreach ($i  in $metadata) {

        $message_metadata = $i['metadata']
        $message_urls = $i['links']

        #TODO FIX ; got error message error casting to system.uri for [hashtable]
        foreach ($url in $message_urls) {
            # $url[0] is the URL, $url[1] is the text description
            try{
                $u = [system.uri]$url['url']
                $url_host = $u.Host
                }
            
            catch {
                "Error casting url to system.uri: " + [string]$url
            }

            $query_only = $u.AbsoluteUri
            #Strip trailing "/" if present
            if ($query_only[-1] -eq "/") { $query_only = $query_only -replace ".$"}

            #make array of all url hosts (all $i[0]s)
            $url_hosts_present = foreach($i in $uniqueURLs) {$i[0]}

            if (-not $url_hosts_present) {
                
                $uniqueURLs.add(@($url_host, @($query_only, $url['text'], $message_metadata, 1))) > $null

            } elseif ($url_hosts_present.contains($url_host)) {

                #see if we already have listed the full URL
                $urls_present = foreach($i in $uniqueURLs[$uniqueURLs.IndexOf($url_host)]) {$i[0]}

                if (-not $urls_present.contains($query_only)) {
                    #if the url_host is found but not this specific url, add
                    $uniqueURLs[$uniqueURLs.indexOf($url_host)] += (@($query_only, $url['text'], $message_metadata, 1)) > $null
                
                } else {
                    # increment count of messages with URL found
                    $uniqueURLs[$uniqueURLs.IndexOf($url_host)][1][3] += 1
                }

            } else {

                $uniqueURLs.add(@($url_host, @($query_only, $url['text'], $message_metadata, 1))) > $null

            }
        }
    }
}


function Get-HTMLContent($data) {
    #return all characters between <body and </body tags as single string.

    return $data[$data.indexof("<body")..$data.indexof("</body")] -join ""
 }

 function Get-Metadata($data, $url_pattern, $urlfilters) {
    <#
     * walk path and uncompress and .zip files
     * open any .msg files, extract URLS and email metadata
     * export results as csv

     :param: data  {boolean} Ignore emails that were found in Junk Email folders
     :param: url_pattern     {string}  Directory containing emails or .zips of emails
     :return: {array}  array of (metadata, urls)
     #>

     $date_pattern = "\bDate:\s(?<date>[\w,: ]*)"
     $sender_pattern = "\bFrom:\s(?<sender>[,\`"\w @<>\.\]\[\-_]*)"
     $sender_ip_pattern = "\bsender\sip\sis\s([^\)]*)"
     $recipient_pattern = "\bTo:\s(?<recipient>[,\`"\w @<>\.\[\]\-_]*)"

     
     $message_urls = New-Object System.Collections.ArrayList
     $message_metadata =@{}

     $html_data = Get-HTMLContent -data $data
     $extracted_urls = Select-String -InputObject $html_data -Pattern $url_pattern -AllMatches 

     :outer foreach ($i  in $extracted_urls.Matches) { 

        #do not add duplicate entries to $message_urls
        foreach ($j in $message_urls) {
            if ($i.groups['url'].value -eq ($j['url'])) {
                break outer
            }            
        }

        foreach ($k in $urlfilters) {
            if ($i.groups['url'].value.contains($k)) {
                break outer
            }
        }

        $message_urls += ,@{url = $i.groups['url'].value; text = $i.groups['text'].value }                
    }

    # Remove ascii 10 or ascii 13 from text and url
    foreach ($i in $message_urls) {
        if ($i['text'].contains([char]13) -or $i['text'].contains([char]10)) {
            $i['text'] = $i['text'] -replace [char]13,'' -replace [char]10,'' 
        }
        if ($i['url'].contains([char]13) -or $i['url'].contains([char]10)) {
            $i['url'] = $i['url'] -replace [char]13,'' -replace [char]10,'' 
        }        
    }

    # | out-null required otherwise result will append to return value for entire function call
    # see https://stackoverflow.com/questions/8671602/problems-returning-hashtable
    $data.replace("\u003e", ">") |out-null 
    $data.replace("\u003c", "<")|out-null  


    $extracted_date = Select-String -InputObject $data -Pattern $date_pattern -AllMatches
    $extracted_sender = Select-String -InputObject $data -Pattern $sender_pattern -AllMatches
    $extracted_sender_ip = Select-String -InputObject $data -Pattern $sender_ip_pattern -AllMatches
    $extracted_recipient = Select-String -InputObject $data -Pattern $recipient_pattern -AllMatches
    $extracted_subject = $file.Name -replace ".{4}$" # Default filename is subject + extension of .msg

    if (-not $extracted_recipient) {  $extracted_recipient = "Null"} else {$extracted_recipient = $extracted_recipient.Matches.groups[-1].Value}
    if (-not $extracted_sender) { $extracted_sender = "Null"} else { $extracted_sender = $extracted_sender.Matches.groups[-1].Value}
    if (-not $extracted_sender_ip) { $extracted_sender_ip = "Null"} else { $extracted_sender_ip = $extracted_sender_ip.Matches.groups[1].Value}
    if (-not $extracted_date) { $extracted_date = "Null"} else { $extracted_date = $extracted_date.matches.groups[-1].value}  


    if ($extracted_date.Length -ge 100 -or $extracted_recipient.Length -ge 100 -or $extracted_sender.length -ge 110 -or $extracted_sender_ip.length -ge 100) {
         "regex failed"
    }

    if ($extracted_date -and $extracted_date -ne "Null") {
        $extracted_date = Format-DateTime -string $extracted_date 
    }  else {
        "date regex failed"
    }
     
     $message_metadata.subject = $extracted_subject    
     $message_metadata.sender = $extracted_sender
     $message_metadata.date =$extracted_date
     $message_metadata.recipient = $extracted_recipient
     $message_metadata.sender_ip = $extracted_sender_ip

    foreach ($i in $message_metadata.Values) {
        if ($i -contains "\r" -or $i -contains "\n") {
            $i.replace("[`r`n\r\n]*","") | Out-Null
        } if ($i -contains "\u003c") {
            $i.replace("\u003c", "<") | Out-Null
        } if ($i -contains "\u003e") {
            $i.replace("\u003e", ">") | Out-Null
        }

    }

    return @{metadata = $message_metadata;links=$message_urls}

 }



if ($URLFilterList) {
    $URLFilters = get-content -Path $URLFilterList
}

if (-not $writeCSV -and -not $writeJSON) {
    "No output format specified.  Please specify at least one of -writeCSV or -writeJSON"    
    [System.Environment]::exit(1)
} elseif (-not $path) {
    "No path specified.  Please use -path followed by path to .msg files.  Path will be recursively searched."
    [System.Environment]::exit(1)
}

main -path $path -FilterJunkFolders $FilterJunkFolders -urlfilters $URLFilters -writeCSV $writeCSV -writeJSON $writeJSON -verboseOutput $verboseOutput