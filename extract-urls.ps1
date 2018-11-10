param(
    [switch] $FilterJunkFolders,
    [String] $path)
function main($FilterJunkFolders, $path) {
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
    #$path = "c:\out\_run3\"
    "Scanning " + $path
    $files = Get-ChildItem -path $path -file -recurse

    $zips = @()
    foreach ($file in $files)    {
        if ($file.Extension -eq ".zip") {
            $zips += $file }
    }

    [String]$zips.Length + " zip files found"
     
    #TODO uncompress zip files
    
    [String]$files.Length + " files found"

    foreach ($file in $files) {
        
        $completion = [math]::Round($files.indexOf($file)/$files.Length * 100)
        $str = "Parsing " + [String]$files.Length +" Message Files"
        Write-Progress -Activity $str -Status "$completion% Complete" -PercentComplete $completion        

        if ($file.Extension -eq ".msg")  {
            #Only parse .msg files
            try {
                $data = [io.file]::ReadAllText($file.FullName) }
            catch {
                $errors.add($file.FullName) > Null
                continue}

            if ($data -match $url_pattern) {

                $metadata.add(@(Get-Metadata($data, $url_pattern))) > $null #append results but suppress output 
            }
        }
    }


    $stream = [System.IO.StreamWriter]::new($path + "data.csv")
    $stream.writeline("Subject, Recipient, Sender, Sender IP, Date, URLs")
    
    $metadata| ForEach-Object {
        
        $stream.WriteLine( $_[0]+ $_[1] -join ", ")}
    "Wrote " + $path + "data.csv"
    $stream.Close()
    
    $stream = [System.IO.StreamWriter]::new($path + "errors.csv")
    $stream.writeline("Errors")
    
    $errors| ForEach-Object {
        
        $stream.WriteLine( $_)}
    "Wrote " + $path + "errors.csv"
    $stream.Close()

    [String]$metadata.Count + " messages parsed successfully."
    [String]$errors.Count + " messages could not be opened."
    $endtime = get-date
    "Execution ended at: " + $endtime
    "Execution Duration: {0:HH:mm:ss}" -f ([datetime]($endtime - $starttime).Ticks) 
    

 }

 function Get-Metadata($data, $url_pattern) {
    <#
     * walk path and uncompress and .zip files
     * open any .msg files, extract URLS and email metadata
     * export results as csv

     :param: data  {boolean} Ignore emails that were found in Junk Email folders
     :param: url_pattern     {string}  Directory containing emails or .zips of emails
     :return: {array}  array of (metadata, urls)
     #>

     $url_pattern = "\b([a-zA-Z]{3,})://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)*?"
     $date_pattern = "\bDate:\s([^+]*\s)"
     $sender_pattern = "\bFrom:\s([^>]*>)"
     $sender_ip_pattern = "\bsender\sip\sis\s([^\)]*)"
     $recipient_pattern = "\bTo:\s([^>]*>)"

     $extracted = Select-String -InputObject $data -Pattern $url_pattern -AllMatches
     $message_urls = @()
     $message_metadata =@()

     foreach ($i  in $extracted.Matches) { 
         $message_urls += $i.Value }

     $extracted_date = Select-String -InputObject $data -Pattern $date_pattern -AllMatches
     $extracted_sender = Select-String -InputObject $data -Pattern $sender_pattern -AllMatches
     $extracted_sender_ip = Select-String -InputObject $data -Pattern $sender_ip_pattern -AllMatches
     $extracted_recipient = Select-String -InputObject $data -Pattern $recipient_pattern -AllMatches
     $extracted_subject = $file.Name -replace ".{4}$" # Default filename is subject + extension of .msg

     #$temp = @()

     if (-not $extracted_recipient) {  $extracted_recipient = "Null"} else {$extracted_recipient = $extracted_recipient.Matches[-1].Value}
     if (-not $extracted_sender) { $extracted_sender = "Null"} else { $extracted_sender = $extracted_sender.Matches[-1].Value}
     if (-not $extracted_sender_ip) { $extracted_sender_ip = "Null"} else { $extracted_sender_ip = $extracted_sender_ip.Matches[-1].Value}
     if (-not $extracted_date) { $extracted_date = "Null"} else { $extracted_date = $extracted_date.Matches[-1].Value}

     $message_metadata = $extracted_subject, $extracted_recipient, $extracted_sender, $extracted_sender_ip, $extracted_date
     #$temp = @($message_metadata, $message_urls)
     return @($message_metadata, $message_urls)

 }
"path is set to " + $path
main -path $path -FilterJunkFolders $FilterJunkFolders