# Version TPS-0917Jan27
[CmdletBinding()]
param(
    [String]$Value
)


# the following two lines set up an alias to allow the user to bring up the IP Address Table that is being used.
# once the table pops up, they can query the contents, which will then be output to the console 
function get-IPAddressTable {$IPAddressTable | Out-GridView -Title "IP Address Table" -OutputMode Multiple | ft -AutoSize}
set-alias IPT get-IPAddressTable

cls
# cakk set-alternatingRows
. C:\TPSITS\out-consolegraph.ps1
 & C:\TPSITS\set-AlternatingRows.ps1 
. 'C:\TPSITS\set-AlternatingRows.ps1' 
#& 'C:\Program Files\WindowsPowerShell\Scripts\SetAlternatingRows.ps1'
#. 'C:\Program Files\WindowsPowerShell\Scripts\SetAlternatingRows.ps1'

if ($PSVersionTable.PSVersion.Major -lt 5)
{
    Write-Host "IP Analyzer not tested on this version of PowerShell. Recommend version 5."  -ForegroundColor DarkCyan
}


function send-TPSEmail {

$smtpServer = "smtplb.torontopolice.on.ca" 
$msg = new-object Net.mail.mailmessage
$msg.IsBodyHTML = $true


# add attachments
$attachments   = Get-ChildItem $mypath\* -Include *.html, *.xlsx
foreach ($att in $attachments) {

$atts = new-object Net.Mail.Attachment($att) 
$msg.Attachments.Add($atts)
}



#$message.Body = $message.Body = ps | Select-Object Name, id, handles | ConvertTo-Html -Head $header

$smtp = new-object Net.Mail.SmtpClient($smtpServer)

$msg.From = "Automatron@torontopolice.on.ca" 
#currently logged on user
[String[]]$Recipient = (Get-ADUser $env:USERNAME -properties mail).mail
$Recipient | % {$msg.to.add($_)}
$msg.subject =  "IP Analysis results for $env:USERNAME on $(get-date -Format f)."
$msg.body = $RT

# work around block {open, send, close}
Stop-Service mcshield
Start-Sleep -s 1
$smtp.Send($msg)
Start-service mcshield -WarningAction SilentlyContinue

}





Function Check-Authorization {
# the following code is used to provide cursory authorization
# only members of the groups included in the $groups are permitted to use this application unless
# $byPassAuthentication has been enabled.  THis is meant to be used just for testing or for members
# working outside of TPS AD

$byPassAuthentication = $true
$user = $env:username
#$user = "b10665"
$groups = "cus","cfa","Tsg"
$authMembers = @()

# is computer part of TPS domain
if ($env:USERDOMAIN -eq "PRD")
{

foreach ($group in $groups)
{
$members = Get-ADGroupMember -Identity $group -Recursive | Select -ExpandProperty Name -ErrorAction SilentlyContinue
$authMembers += $members
}
}

If ($authMembers -contains $user) {
      Write-Host "User $user must exist in one or more of the following groups: $groups."
      Write-Host "$user authorized to use this application." -ForegroundColor green
      
 } Else {
        Write-Host "User $user not found in the following groups: $groups or not on tps.prd network."
        if ($byPassAuthentication -eq $true){Write-Host "Authorization bypass enabled. $user bypassing authorization ....."`n -ForegroundColor Green }
        else {
        Write-Host "$user not authorized. Please talk to your system administrator." -ForegroundColor Red
        exit}
}

  
}

Write-Host "Welcome to TPS IP Production Order Utility .... $(get-date)"`n
Check-Authorization


function IP-Mapper ($parsedFileSummary, $file, $path)
{

$TableInfo = $null


#clean out files from previous runs of this script
#remove-item -Path $myPath\*.html
#remove-item -Path $myPath\*.xlsx

#$Path = Split-path $MyInvocation.MyCommand.Definition

$Locations = ForEach ($Addr in $parsedFileSummary )
{ 
  

    [PSCustomObject]@{
        Qty = $Addr.Count
        Lat = $Addr.Lat
        Lng = $Addr.Lon
        IP = $Addr.IPAddress
          }
}

$Markers = "var markers = [`n"
# add the Vanity Registration IP as the first item in the list
$Markers += "    ['666','$RegistrationAddress',43.7,-79.7]"
$Markers += ",`n"

ForEach ($Num in (0..($Locations.Count - 1)))
{   $Location = $Locations[$Num].IP
    $Markers += "    ['$($Locations[$Num].Qty)','$Location',$($Locations[$Num].Lat),$($Locations[$Num].Lng)]"
    If ($Num -lt ($Locations.Count - 1))
    {   $Markers += ",`n"
    }
}
$Markers += "`n  ];`n"

 $TableInfo = $null
 $TableInfo += "<li><p>$RegistrationAddress<br>Qty: 666<br>Lat: 43.7<br>Lon: -79.7</p></li>"

ForEach ($Location in $Locations)
{   $TableInfo += "<li><p>$($Location.IP)<br>Qty: $($Location.Qty)<br>Lat: $($Location.Lat)<br>Lon: $($Location.Lng)</p></li>"
}


$HTML = @"
<!DOCTYPE html>
<html>
  <head>
    <meta name="viewport" content="initial-scale=1.0, user-scalable=no">
    <meta charset="utf-8">
    <title>IP map for $vanity</title>
    <h3>IP Map for $vanity  --- IP address: $RegistrationAddress</h3>
    <hr>
    <style>
      html, body {
        margin: 0px;
        padding: 0px;
        background-color:lightblue
     }
     h3 {
    text-indent: 500px;
    }
      #map-canvas {
        height: 900px;
        width: 1350px;
        margin: 0px;
        padding: 0px
      }
      td {
        vertical-align: top;
      }
    </style>
    <script src="https://maps.googleapis.com/maps/api/js?v=3.exp&sensor=false"></script>
    <script>
function initialize() {
  $Markers
  var bounds = new google.maps.LatLngBounds();
  var mapOptions = {
    mapTypeId: google.maps.MapTypeId.HYBRID,
  }
  var map = new google.maps.Map(document.getElementById('map-canvas'), mapOptions);
  
  for( i = 0; i < markers.length; i++ ) {
    var position = new google.maps.LatLng(markers[i][2], markers[i][3]);
    bounds.extend(position);
    var image = 'http://chart.apis.google.com/chart?chst=d_map_pin_letter&chld=' + (i + 1) + '|FF0000|000000';
    marker = new google.maps.Marker({
      position: position,
      map: map,
      title: markers[i][0] + '\n' + markers[i][1],
      icon: image
    });
  }

  if (markers.length > 1) {
    map.fitBounds(bounds);
  }
  else {
    map.setCenter(new google.maps.LatLng(markers[0][1],markers[0][2]));
    map.setZoom(7);
  
  }
}

google.maps.event.addDomListener(window, 'load', initialize);
    </script>
  </head>
  <body>
    <table>
      <td id="map-canvas"></td>
     <td><b><ol start="0">$tableInfo</ol></b></td>
    </table>
    </body>
</html>
"@

Write-Verbose "$(Get-Date): Saving file..."
Try {
    $HTML | Out-File $myPath\$file.html -Encoding ASCII -ErrorAction Stop
}
Catch {
    Write-Warning "Unable to save HTML at $myPath\$file.html because $($Error[0])"
    Exit
}

# uncomment the following if you want to have the  browser open up the map with all of the points
#& $myPath\$file.html

#$RT = $ParsedFile | group IPAddress | Select 'Count','Name' | sort -Descending count | ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-String
#$rt =  gc C:\ooops2.htm | Out-String
#Send-YahooMail "rturnbul@yahoo.com" "DrDConway69!" "Rturnbul@yahoo.com" "SuperCool" $RT

} 




function Send-YahooMail ($Username, $Password, $RcptTo, $Subject, $Body, $attachments   ){
    $SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
    $Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $SecurePassword
    Read-Host 
    Send-MailMessage -From $Username -To $RcptTo -Subject $Subject -Body $Body  -BodyAsHtml -SmtpServer smtp.mail.yahoo.com -Port 587 -UseSsl -Credential $Credentials -Attachments $attachments
   
   
    }

function Get-ReferencesFromPdf
{
<#
    .SYNOPSIS
        Template script
    .DESCRIPTION
        This script sets up the basic framework that I use for all my scripts.
    .PARAMETER
    .EXAMPLE
    .NOTES
        ScriptName : 
        Created By : Ron
        Date Coded : 12/17/2016 08:02:14
        
        ErrorCodes
            100 = Success
            101 = Error
            102 = Warning
            104 = Information
    .LINK
        https://code.google.com/p/mod-posh/wiki/Production/
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $Path
    )

    $Path = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)

    try
    {
        $reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $Path
    }
    catch
    {
        throw
    }

   
    $number = ''

    for ($page = 1; $page -le $reader.NumberOfPages; $page++)
    { 
        $lines = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $page) -split "\r?\n"
 
 $registrationAddressFound = $false
 $UTCTimeFound = $false
        
        
        foreach ($line in $lines)
        {
            #write-host "This line contains $line " -backgroundColor blue
            $UTCTimeFound = $false
            
            switch -Regex ($line)
            {

             '^Service'
            {
            $serviceName = $line.TrimStart("Service ")
            #write-host "Service name is $serviceName"   -backgroundColor green
            } 

            
            '^Vanity Name'
            {
            $vanityName = $line.TrimStart("Vanity Name ")
            #write-host "Vanity name is $vanityName"   -backgroundColor Magenta
            } 

            '^Registration IP'
            {
            $registrationIP = $line.TrimStart("Registration Ip ")
            #write-host "Registration IP is $registrationIP"   -backgroundColor Cyan
            break
            } 
            
            '(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))'
                {
                
                  $IPAddress = $line  -replace "IP Address "
                  If ($IPAddress -match "IP Addresses")
                  { $IPAddress = $IPAddress -replace "IP Addresses "}
                  
                #write-host "$IPAddress is a IPV6 address"
                  $IPAddressFound = $true
                 $previousAddress = $line
                }



             '^Time \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} UTC'
             
                {
        
                
                 $UTCTime =  ($line.TrimStart("Time ")).TrimEnd(" UTC") 
                 #write-host "$UTCTime is a UTC time"
                 #$UTCTime =  $temp
                 $UTCTimeFound = $true
                 }

              



                '\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b'
              
                {
                    
                  $IPAddress = $line -replace "IP Address "
                  If ($IPAddress -match "IP Addresses")
                  { $IPAddress = $IPAddress -replace "IP Addresses "}
                  
                #write-host "$IPAddress is a IPV4 address"
               $IPAddressFound = $true
                 
                            #break

                }

            }
            
                       
                if ( $UTCTimeFound -eq $true)
                {
                New-Object psobject -Property @{
                        RegistrationAddress = $registrationIP
                        VanityName = $vanityName
                        IPAddress = $IPAddress
                        UTCTime = $UTCTime
                        File = ($file.DirectoryName + "\" + $file.Name)
                        <# build the entire object out for the IP lookup step
                        
                        as = ""
                        city = ""
                        country = ""
                        countryCode = ""
                        isp = ""
                        lat = ""
                        lon =  ""
                        org = ""
                        #query = ""
                        region = ""
                        regionName = ""
                        status = ""
                        
                        zip = ""
                        #>
                    }
                     
                    }
                    
        } #$previousLine = $line; #write-host "Previous line is $previousLine"
    }

    $reader.Close()
}

$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: lightblue;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
</style>
<title>
IP Traffic File for $vanity
</title>
"@

$IPTableHeader = @"
<style>

TABLE {font-size: xx-small;border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {font-size: xx-small;border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color:lightgreen;}
TD {font-size: xx-small;border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }

</style>
<title>
IP Table Details 
</title>
"@


 

#Unblock-File -Path C:\Scripts\PdfToText\iTextSharp.dll
Add-Type -Path C:\itextsharp.dll
#cd  C:\Users\Ron\Downloads\AlsonTest

# add code to verify if current user is member of authorized group


#get current user
$currentUser = (Get-WmiObject  –Class Win32_ComputerSystem | Select-Object UserName).UserName.Split("\")[1]
#$currentUser = $currentuser.UserName.Split("\")[1]

# location where current user will save all of the documents that are created
# each user has their own location; each run of the tool will save files in a different location
$myPath = [Environment]::GetFolderPath('MyDocuments')+"\Automatron" 
# + (get-date).Ticks

remove-item -Path $myPath\*.html
remove-item -Path $myPath\*.xlsx


$Path = Split-path $MyInvocation.MyCommand.Definition

ri .\output.txt -ErrorAction SilentlyContinue

Write-host "A log of your transactions is kept in the Output.txt file."
Write-host "Type 'gc $mypath\output.txt' to view the contents."`n
New-Item -ItemType file output.txt -ErrorAction SilentlyContinue

# nullify session variables; initialize variables
$IPAddressTable = $null
$IPAddressTableFile = "$mypath\IPAddressTableFile.csv"
$parsedFile = @()
$IPAddressTable = @()
$newIP = @()
$parsedFile = $null
$attachments = @()
$files = @()

#load IP Address table from disk
# this table maintains a running list of all IP addresses that have ever been gathered by this tool
# this allows for additional processing of this list plus significantly reduced the number of times that an IP address needs to be geocoded

if(test-path $IPAddressTableFile)
{#load table into an array so that we can do lookups against it
$IPAddressTable = import-csv $IPAddressTableFile

Write-Host "`n`nUsing IP Address Table File from $IPAddressTableFile containing $($IPAddressTable.count) IP record(s)."

#start-sleep -Seconds 5
# show the list of IPs to the user
#$IPAddressTable | Out-GridView -Title "IP Address Table" -OutputMode Multiple | ft -AutoSize  
}
else
{<# make a new IPAddressTable object
 THis code will add a new IP address to the IPAddressTable
$newIPHash = @{

as=""              
city=""                                  
country=""                                 
countryCode=""                                  
IPAddress=""
isp=""                        
lat=""                                     
lon=""                                  
org=""                       
region=""                                       
regionName=""                              
zip="" 
}                                 
    $IPAddressTable += $IPAddressTable +   (New-Object PSObject -Property $newIPHash)
    #>
}


$rt= $null

$filesToSearch = gci $myPath *.pdf | Out-GridView -Title "Input Files List"  -OutputMode Multiple
$myTime = (get-date).Date
$rt = "The following $($filesToSearch.count) file(s) were analyzed on $(Get-Date -UFormat %D) at $(Get-Date -UFormat %T): <br><br>" 
foreach ($file in $filesToSearch) {$rt= $rt + "---------->  " + $file.name + "<br><br>"| Out-String }
$IPAddressFound = $false

foreach ($file in $filesToSearch) {
 #Note - Is there a benefit to keep track of all files that have ever been run; determine if there is overlap

 $parsedFile = $null
$parsedFile += $currentfile = Get-ReferencesFromPdf ($file.DirectoryName + "\" + $file.Name)


$parsedFile >> $mypath\output.txt 
$vanity = $parsedFile.vanityname[0]
$RegistrationAddress = $parsedFile.registrationAddress[0]

Write-host "`n`n"
$parsedFile | Group-Object -Property ipaddress | sort -Property count -Descending | Out-ConsoleGraph -Property count -Title "IP Analysis for $vanity" -HighColor red -med Yellow -LowColor green

$out = $myPath + ($file.name)+ ".csv"
$Pre = "`n$vanity - $registrationAddress "
$Post = "<P>$File</P><P>Automatron.</p><br><br>"
$parsedFile | group IPAddress | Select 'Count','Name' | sort -Descending count | ConvertTo-HTML -Head $Header -PreContent $Pre -PostContent $Post | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd > $out

# create add Lat, Lon fields using select-object since we will need to geocode using local IPAddressTable or using online search
$parsedFileSummary = $parsedFile | group IPAddress  | Select 'Count',@{Name="IPAddress";Expression={$_.Name}}, Lat, Lon | sort -Descending count 



$numLoops = $parsedFileSummary.count-1

for ($i=0; $i -le $numLoops ; $i++)
 {
# Do a lookup for each IP in ParsedFileSummary  to see if it already exists in IPAddressTable, if it does get Lat and Lon.  
# If it doesn't IP address is new and this address needs to be added to IPAddressTAble and also to this object $parsedFileSummary 
If ($IPAddressTable.IpAddress -contains $parsedFileSummary[$i].IPAddress)
{#Write-Host "Found this IP Address in the IP Address table ...."   

 

$myIndex = (0..($IPAddressTable.Count-1) | where {$IPAddressTable[$_].ipaddress -eq $parsedFileSummary[$i].IPAddress})

$parsedFileSummary[$i].lat = $IPAddressTable[$myIndex].lat
$parsedFileSummary[$i].lon = $IPAddressTable[$myIndex].lon



}
else
{# If it doesn't, IP address is new and this address needs to be added to IPAddressTAble and also to this object $parsedFileSummary 


$request = "http://ip-api.com/json/$($parsedFileSummary[$i].IPAddress)"
#Write-Host "This is number $i out of $numLoops item."
$getIP = Invoke-WebRequest $request | ConvertFrom-Json

$newIPHash = @{IPAddress = $parsedFileSummary[$i].IPAddress
AS = $getIP.as
CITY =  $getIP.CITY
COUNTRY = $getIP.COUNTRY
COUNTRYCODE = $getIP.countryCode
ISP = $getIP.isp
LAT = $getIP.lat
LON = $getIP.lon
ORG = $getIP.org
REGION = $getIP.region
REGIONNAME = $getIP.regionName
ZIP = $getIP.zip }

$newIP = New-Object psobject -Property $newIPHash

#write-host "Adding new record ....$newIP"

# add fields to the $parsedFileSummary Table
$parsedFileSummary[$i].Lon = $getIP.lon
$parsedFileSummary[$i].Lat = $getIP.lat

$IPAddressTable += $newIP                
}

}

#write the  $IPAddressFile object to disk for importing at start of next run of this script
$IPAddressTable | Export-Csv -Path $mypath\IPAddressTableFile.csv -NoTypeInformation

#map all of the conversations from each source file
IP-Mapper $parsedFileSummary $file.Name $myPath


 Import-Module -FullyQualifiedName "C:\Program Files\WindowsPowerShell\Modules\ImportExcel"

 $parsedFile = $parsedFile | select * | sort-object -Property UTCTIME -Unique
 
 <#
 ri .\ronnytease*

 $parsedFile   | Export-Excel ronnytease9.xlsx -Show -AutoSize `
        -IncludePivotTable `
        -IncludePivotChart `
        -ChartType ColumnClustered `
        -PivotRows isp,IPAddress `
        -PivotData @{Number='sum';File='count'}

        
  
   #>


  # the next lines are responsible for exporting the files to Excel and creating a pivot table/chart
# multiple table/chart combinations can be made during a single run of the code, as shown below
# the results can be placed on their own worksheet using -worksheetname parameter

Get-Process excel -ErrorAction SilentlyContinue | stop-process –force -ErrorAction SilentlyContinue | Out-Null

# need to remove the file at the beginning of each run
$IPspreadSheet = $myPath + "\$file.xlsx"
#Start-Sleep -Seconds 5
#ri $IPspreadSheet  -Force     -ErrorAction Ignore


$parsedFile | where IPAddress | select UTCTime, VanityName, File, IPAddress, RegistrationAddress|
    Export-Excel $IPspreadSheet  -AutoSize -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname $vanity `
        -IncludePivotTable `
        -IncludePivotChart `
        -ChartType PieExploded3D `
        -PivotRows IPAddress,UTCTime `
        -PivotData @{IPAddress='count'} `
        -PivotDataToColumn

$rt =  $rt + (gc $out | Out-String) 

}

$IPTable = ($ipaddresstable | sort -Descending -Property region    | ConvertTo-HTML -Head $IPTableHeader | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd)


$rt = $rt + $IPTable

Get-Process excel -ErrorAction SilentlyContinue | stop-process –force -ErrorAction SilentlyContinue | Out-Null
Start-Sleep -Seconds 5
$attachments = gci $myPath\* -Include *.html, *.xlsx
$date = (get-date).DateTime

#determine whether mail is being sent from inside TPS or from outside

if($env:USERDOMAIN -eq "PRD") {
send-TPSEmail}
else {
Send-YahooMail "rturnbul@yahoo.com" "DrDConway69!" "Rturnbul@gmail.com" "Your IP Analysis results for $date " $RT -attachments $attachments 
}

If ($error[0] -Notmatch "Unable")
   {Write-Host "Message sent successfully!" -ForegroundColor Green}
   else
    {Write-Host "Message not sent! Call your administrator or try again." -ForegroundColor Red}