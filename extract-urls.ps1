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

    "Scanning " + $path
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

    [String]$zips.Length + " zip files found"
     
    #TODO uncompress zip files
    
    [String]$files.Length + " files found"
    [String]$msgFiles.Count + " message files found"

    foreach ($file in $msgfiles) {
        
        $completion = [math]::Round($msgfiles.indexOf($file)/$files.Length * 100)
        $str = "Parsing " + [String]$msgfiles.Count +" Message Files"
        Write-Progress -Activity $str -Status "$completion% Complete" -PercentComplete $completion        
          
        try {
            $data = [io.file]::ReadAllText($file.FullName) }
        catch {
            $errors.add($file.FullName) > Null
            continue}

        if ($data -match $url_pattern) {

            $metadata.add(@(Get-Metadata -data $data -url_pattern $url_pattern -urlfilters $urlfilters)) > $null #append results but suppress output 
        }
        
    }

    if ($metadata.Count -gt 0) {
        Format-UniqueMetadata -metadata $metadata -uniqueURLs $uniqueURLs  
        
        $stream = [System.IO.StreamWriter]::new($path + "\unique.csv")
        $stream.writeline("URL, Domain, Subject, Recipient, Sender, Sender IP, Date, additional recipients")
        
        foreach ($i in $uniqueURLs | Sort-Object) {
            $len = $i[1][1].Count -1
            $meta = $i[1][1][0..$len] -join ";"
            $str = $i[1][0], $i[0] -join " , "
            $str = $str, $meta, $i[1][2] -join ";"

            $stream.WriteLine($str)
        }
        "Wrote " + $path + "\unique.csv with " + $uniqueURLs.Count + " unique URLs"
        $stream.Close()
    }

    $stream = [System.IO.StreamWriter]::new($path + "\data.csv")
    $stream.writeline("Subject, Recipient, Sender, Sender IP, Date, URLs")
    
    $metadata| Sort-Object |ForEach-Object {
        
        $stream.WriteLine( $_[0]+ $_[1] -join ", ")}
    "Wrote " + $path + "\data.csv"
    $stream.Close()

    if ($errors.Count -gt 0) {
        $stream = [System.IO.StreamWriter]::new($path + "\errors.csv")
        $stream.writeline("Errors")
        
        $errors| Sort-Object |ForEach-Object {
            
            $stream.WriteLine( $_)}
        "Wrote " + $path + "\errors.csv"
        $stream.Close()
    }

    [String]$metadata.Count + " messages parsed successfully."
    [String]($msgFiles.Count - $metadata.Count) + " messages had no URL"
    [String]$errors.Count + " messages could not be opened."
    $endtime = get-date
    "Execution ended at: " + $endtime
    "Execution Duration: {0:HH:mm:ss}" -f ([datetime]($endtime - $starttime).Ticks) 
    

 }

function Format-UniqueMetadata($metadata, $uniqueURLs) {
    <#
    * reformat $metadata into list of unique urls instead of one result per message

    :param: $metadta
    :param: $uniqueURLs

    #>

    foreach ($i  in $metadata) {

        $message_metadata = $i[0]
        $message_urls = $i[1]

        foreach ($url in $message_urls) {
            $u = [system.uri]$url
            $url_host = $u.Host

            #if ($u.Query) { $query_only = $u.AbsoluteUri.Replace($u.Query, "") 
            #} else {        $query_only = $u.AbsoluteUri
            #}
            $query_only = $u.AbsoluteUri
            #Strip trailing "/" if present
            if ($query_only[-1] -eq "/") { $query_only = $query_only -replace ".$"}

            #make array of all url hosts (all $i[0]s)
            $url_hosts_present = foreach($i in $uniqueURLs) {$i[0]}

            if (-not $url_hosts_present) {
                
                $uniqueURLs.add(@($url_host, @($query_only, $message_metadata, 1))) > $null

            } elseif ($url_hosts_present.contains($url_host)) {

                #see if we already have listed the full URL
                $urls_present = foreach($i in $uniqueURLs[$uniqueURLs.IndexOf($url_host)]) {$i[0]}

                if (-not $urls_present.contains($query_only)) {
                    #if the url_host is found but not this specific url, add
                    $uniqueURLs[$uniqueURLs.indexOf($url_host)] += (@($query_only, $message_metadata, 1)) > $null
                
                } else {
                    # increment count of messages with URL found
                    $uniqueURLs[$uniqueURLs.IndexOf($url_host)][1][2] += 1
                }

            } else {

                $uniqueURLs.add(@($url_host, @($query_only, $message_metadata, 1))) > $null

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

     $url_pattern = "\b([a-zA-Z]{3,})://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)*?"
     $date_pattern = "\bDate:\s(?<date>[\w,: ]*)"
     $sender_pattern = "\bFrom:\s(?<sender>[\w @<>\.\]\[]*)"
     $sender_ip_pattern = "\bsender\sip\sis\s([^\)]*)"
     $recipient_pattern = "\bTo:\s(?<recipient>[\w @<>\.\[\]]*)"

     
     $message_urls = @()
     $message_metadata =@()
     #TODO urlpattern regex needs fixing...multiple matches and doesnt get the href text.

     $html_data = Get-HTMLContent -data $data
     $extracted_urls = Select-String -InputObject $html_data -Pattern $url_pattern -AllMatches

     foreach ($i  in $extracted_urls.Matches) { 
         if ((-not $message_urls.Contains($i.Value) -and (-not $urlfilters.contains($i.Value)))) {
            $message_urls += $i.Value }}

     $extracted_date = Select-String -InputObject $data -Pattern $date_pattern -AllMatches
     $extracted_sender = Select-String -InputObject $data -Pattern $sender_pattern -AllMatches
     $extracted_sender_ip = Select-String -InputObject $data -Pattern $sender_ip_pattern -AllMatches
     $extracted_recipient = Select-String -InputObject $data -Pattern $recipient_pattern -AllMatches
     $extracted_subject = $file.Name -replace ".{4}$" # Default filename is subject + extension of .msg

     #$temp = @()

     if (-not $extracted_recipient) {  $extracted_recipient = "Null"} else {$extracted_recipient = $extracted_recipient.Matches.groups[-1].Value}
     if (-not $extracted_sender) { $extracted_sender = "Null"} else { $extracted_sender = $extracted_sender.Matches.groups[-1].Value}
     if (-not $extracted_sender_ip) { $extracted_sender_ip = "Null"} else { $extracted_sender_ip = $extracted_sender_ip.Matches.groups[1].Value}
     if (-not $extracted_date) { $extracted_date = "Null"} else { $extracted_date = $extracted_date.matches.groups[-1].value}
    
     if ($extracted_date.Length -ge 100 -or $extracted_recipient.Length -ge 100 -or $extracted_sender.length -ge 110 -or $extracted_sender_ip.length -ge 100) {
         "regex failed"
     }

     $message_metadata = $extracted_subject, $extracted_recipient, $extracted_sender, $extracted_sender_ip, $extracted_date
     #$temp = @($message_metadata, $message_urls)
     return @($message_metadata, $message_urls)

 }

if ($URLFilterList) {
    $URLFilters = get-content -Path $URLFilterList
}

main -path $path -FilterJunkFolders $FilterJunkFolders -urlfilters $URLFilters