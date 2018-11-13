param(
    [switch] $FilterJunkFolders,
    [String] $path,
    [string] $URLFilterList ="urlfilters.conf")
function main($FilterJunkFolders, $path, $urlfilters) {
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

    $zips = @()
    foreach ($file in $files)    {
        if ($file.Extension -eq ".zip") {
            $zips += $file }
    }

    foreach ($file in $files) {
        if ($file.Extension -eq ".msg") {
            $msgFiles.add($file) > $null
        }
    }

     
    #TODO uncompress zip files

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
            #$metadata.add(@(Get-Metadata -data $data -url_pattern $url_pattern -urlfilters $urlfilters)) > $null #append results but suppress output 
            $temp2 = @(Get-Metadata -data $data -url_pattern $url_pattern -urlfilters $urlfilters)
            $metadata.add($temp2) > $null #append results but suppress output 
        }
        
    }

    if ($metadata.Count -gt 0) {
        Format-UniqueMetadata -metadata $metadata -uniqueURLs $uniqueURLs  
        
        #$stream = [System.IO.StreamWriter]::new($path + "\unique.csv")
        #$stream.writeline("URL, Text, Domain, Subject, Recipient, Sender, Sender_IP, Date, Messages_Received")
        
        <#
        foreach ($i in $uniqueURLs | Sort-Object) {
            $len = $i[1][2].Count -1 #length of metadata ie $i[1]
            $meta = $i[1][2][0..$len] -join ";"
            $str = $i[1][0], $i[1][1], $i[0] -join " , "
            $str = $str, $meta, $i[1][3] -join ";"

            $stream.WriteLine($str)
        }
        #>
        $csv = foreach ($i in $uniqueURLs | sort-object) {
            ConvertTo-Csv -InputObject $i -Delimiter ';'
        }
        $json = foreach ($i in $uniqueURLs | Sort-Object) {
            ConvertTo-Json -InputObject $i -Depth 5
        }
        
        $str = $path + '\' + "unique.csv"
        Export-Csv -InputObject $uniqueURLs -NoTypeInformation -Path $str -Force

        $str = $path + '\' + "unique.json"
        Set-Content -Path $str -Value $json 
        "Wrote " + $path + "\unique.csv with " + $uniqueURLs.Count + " unique URLs"
        "Wrote " + $path +"\unique.json with " + $uniqueURLs.Count + " unique URLs"
        #$stream.Close()
    }


    # write all output as csv

    $csv  = foreach($i in $metadata | sort-object) {
        ConvertTo-Csv -InputObject $i[-1] -Delimiter ';'
    }
    $json = foreach($i in $metadata | Sort-Object) {
        ConvertTo-Json -InputObject $i[-1] -Depth 5
    }
    $str = $path + '\' + "data.csv"   
    $header = "Subject, Recipient, Sender, Sender_IP, Date, URLs"
    set-content -path $str -Value $header
    add-content -path $str -Value $csv
    "Wrote $str"
    $str = $path + '\' + "data.json"
    Set-Content -Path $str -value $json
    "Wrote $str "

    
    $json = foreach ($i in $metadata) {ConvertTo-Json -InputObject $i -depth 5}

    $str = $path + '\' +  "output.json"
    Set-Content -Path $str -Value $json
    

    if ($errors.Count -gt 0) {
        $stream = [System.IO.StreamWriter]::new($path + "\errors.csv")
        $stream.writeline("Errors")
        
        $errors| Sort-Object |ForEach-Object {
            
            $stream.WriteLine( $_)}
        "Wrote " + $path + "\errors.csv"
        $stream.Close()
    }
    "Wrote $str to disk."
    [String]$zips.Length + " zip files found"   
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

        foreach ($url in $message_urls) {
            # $url[0] is the URL, $url[1] is the text description
            try{
                #if ($url.Count -eq 2) {
                #    $u = [system.uri]$url[0]  #got some error messages here, guess url didnt always parse right
                #    $url_host = $u.Host } 
                #else {
                #    $u = [system.uri]$url
                #    $url_host = $u.Host
                $u = [system.uri]$url['url']
                $url_host = $u.Host
                }
            
            catch {
                "Error casting url to system.uri: " + [string]$url
            }

            #if ($u.Query) { $query_only = $u.AbsoluteUri.Replace($u.Query, "") 
            #} else {        $query_only = $u.AbsoluteUri
            #}
            $query_only = $u.AbsoluteUri
            #Strip trailing "/" if present
            if ($query_only[-1] -eq "/") { $query_only = $query_only -replace ".$"}

            #make array of all url hosts (all $i[0]s)
            $url_hosts_present = foreach($i in $uniqueURLs) {$i[0]}

            if (-not $url_hosts_present) {
                
                $uniqueURLs.add(@($url_host, @($query_only, $url[1], $message_metadata, 1))) > $null

            } elseif ($url_hosts_present.contains($url_host)) {

                #see if we already have listed the full URL
                $urls_present = foreach($i in $uniqueURLs[$uniqueURLs.IndexOf($url_host)]) {$i[0]}

                if (-not $urls_present.contains($query_only)) {
                    #if the url_host is found but not this specific url, add
                    $uniqueURLs[$uniqueURLs.indexOf($url_host)] += (@($query_only, $url[1], $message_metadata, 1)) > $null
                
                } else {
                    # increment count of messages with URL found
                    $uniqueURLs[$uniqueURLs.IndexOf($url_host)][1][3] += 1
                }

            } else {

                $uniqueURLs.add(@($url_host, @($query_only, $url[1], $message_metadata, 1))) > $null

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

     #$url_pattern = "\b([a-zA-Z]{3,})://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)*?"
     #$url_pattern2 = @'
#href=\"(?<url>[a-zA-Z]{3,5}:\/\/[^\"]*)\">(?<text>[^(?=<\/a]*)
#'@
     $date_pattern = "\bDate:\s(?<date>[\w,: ]*)"
     $sender_pattern = "\bFrom:\s(?<sender>[,\`"\w @<>\.\]\[]*)"
     $sender_ip_pattern = "\bsender\sip\sis\s([^\)]*)"
     $recipient_pattern = "\bTo:\s(?<recipient>[,\`"\w @<>\.\[\]]*)"

     
     $message_urls = New-Object System.Collections.ArrayList
     $message_metadata =@{}
     #TODO urlpattern regex needs fixing...multiple matches and doesnt get the href text.

     $html_data = Get-HTMLContent -data $data
     $extracted_urls = Select-String -InputObject $html_data -Pattern $url_pattern -AllMatches 

     :outer foreach ($i  in $extracted_urls.Matches) { 

        #do not add duplicate entries to $message_urls
        foreach ($j in $message_urls) {
            if ($i.groups['url'].value -eq ($j['url'])) {
                break :outer
            }            
        }

        foreach ($k in $urlfilters) {
            if ($i.groups['url'].value.contains($k)) {
                break :outer
            }
        }
            #$message_urls += ,@($i.groups[1..2].value) }} }
         $message_urls += ,@{url = $i.groups['url'].value; text = $i.groups['text'].value }                
    }

    $data.replace("\u003e", ">") |out-null #required otherwise function return will have this as well
    $data.replace("\u003c", "<")|out-null  # see https://stackoverflow.com/questions/8671602/problems-returning-hashtable


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
    } 

    else {
        "date regex failed"
    }
     
     $message_metadata.subject = $extracted_subject
     $message_metadata.recipient = $extracted_recipient
     $message_metadata.sender = $extracted_sender
     $message_metadata.sender_ip = $extracted_sender_ip
     $message_metadata.date =$extracted_date

    return @{metadata = $message_metadata;links=$message_urls}

 }

if ($URLFilterList) {
    $URLFilters = get-content -Path $URLFilterList
}

main -path $path -FilterJunkFolders $FilterJunkFolders -urlfilters $URLFilters