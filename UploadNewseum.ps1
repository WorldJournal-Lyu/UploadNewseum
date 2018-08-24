<#

UploadNewseum.ps1

    2018-06-05 Initial Creation

#>

if (!($env:PSModulePath -match 'C:\\PowerShell\\_Modules')) {
    $env:PSModulePath = $env:PSModulePath + ';C:\PowerShell\_Modules\'
}

Get-Module -ListAvailable WorldJournal.* | Remove-Module -Force
Get-Module -ListAvailable WorldJournal.* | Import-Module -Force

$scriptPath = $MyInvocation.MyCommand.Path
$scriptName = (($MyInvocation.MyCommand) -Replace ".ps1")
$hasError   = $false

$newlog     = New-Log -Path $scriptPath -LogFormat yyyyMMdd-HHmmss
$log        = $newlog.FullName
$logPath    = $newlog.Directory

$mailFrom   = (Get-WJEmail -Name noreply).MailAddress
$mailPass   = (Get-WJEmail -Name noreply).Password
$mailTo     = (Get-WJEmail -Name lyu).MailAddress
$mailSbj    = $scriptName
$mailMsg    = ""

$localTemp = "C:\temp\" + $scriptName + "\"
if (!(Test-Path($localTemp))) {New-Item $localTemp -Type Directory | Out-Null}

Write-Log -Verb "LOG START" -Noun $log -Path $log -Type Long -Status Normal
Write-Line -Length 50 -Path $log

###################################################################################





# Set up variables

#$ftp      = Get-WJFTP -Name WorldJournalNewYork
$ftp      = Get-WJFTP -Name Newseum | Where-Object{$_.User -eq 'ny_wjny'}
$ePaper   = Get-WJPath -Name epaper
$workDate = ((Get-Date).AddDays(0)).ToString("yyyyMMdd")
Write-Log -Verb "ftp" -Noun $ftp.Path -Path $log -Type Short -Status Normal
Write-Log -Verb "ePaper" -Noun $ePaper.Path -Path $log -Type Short -Status Normal
Write-Log -Verb "workDate" -Noun $workDate -Path $log -Type Short -Status Normal

Write-Line -Length 50 -Path $log

$pdfName      = "NY_WJNY_" + $workDate + ".pdf"
$uploadFrom   = $ePaper.Path + $workDate + "\optimizeda\" + "NY" + $workDate + "A01.pdf"
$uploadTo     = $ftp.Path + $pdfName
$downloadFrom = $uploadTo
$downloadTo   = $localTemp + (Get-Date).ToString("yyyyMMdd-HHmmss") + ".pdf"
Write-Log -Verb "pdfName" -Noun $pdfName -Path $log -Type Short -Status Normal
Write-Log -Verb "uploadFrom" -Noun $uploadFrom -Path $log -Type Short -Status Normal
Write-Log -Verb "uploadTo" -Noun $uploadTo -Path $log -Type Short -Status Normal
Write-Log -Verb "downloadFrom" -Noun $downloadFrom -Path $log -Type Short -Status Normal
Write-Log -Verb "downloadTo" -Noun $downloadTo -Path $log -Type Short -Status Normal



# Upload file from local to Ftp

$upload = WebClient-UploadFile -Username $ftp.User -Password $ftp.Pass -RemoteFilePath $uploadTo -LocalFilePath $uploadFrom

if($upload.Status -eq "Good"){
    Write-Log -Verb $upload.Verb -Noun $upload.Noun -Path $log -Type Long -Status $upload.Status
}elseif($upload.Status -eq "Bad"){
    $mailMsg = $mailMsg + (Write-Log -Verb $upload.Verb -Noun $upload.Noun -Path $log -Type Long -Status $upload.Status -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $upload.Exception -Path $log -Type Short -Status $upload.Status -Output String) + "`n"
}



# Download file from Ftp to temp from verification

$download = WebClient-DownloadFile -Username $ftp.User -Password $ftp.Pass -RemoteFilePath $downloadFrom -LocalFilePath $downloadTo

if($download.Status -eq "Good"){
    Write-Log -Verb $download.Verb -Noun $download.Noun -Path $log -Type Long -Status $download.Status
}elseif($download.Status -eq "Bad"){
    $mailMsg = $mailMsg + (Write-Log -Verb $download.Verb -Noun $download.Noun -Path $log -Type Long -Status $download.Status -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $download.Exception -Path $log -Type Short -Status $download.Status -Output String) + "`n"
}



# Set hasError status

if( ($upload.Status -eq "Bad") -or ($download.Status -eq "Bad") ){
    $hasError = $true
}





###################################################################################

Write-Line -Length 50 -Path $log

# Delete temp folder

Write-Log -Verb "REMOVE" -Noun $localTemp -Path $log -Type Long -Status Normal
try{
    $temp = $localTemp
    Remove-Item $localTemp -Recurse -Force -ErrorAction Stop
    Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good
}catch{
    $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Bad -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad -Output String) + "`n"
}

Write-Line -Length 50 -Path $log
Write-Log -Verb "LOG END" -Noun $log -Path $log -Type Long -Status Normal
if($hasError){ $mailSbj = "ERROR " + $scriptName }

$emailParam = @{
    From    = $mailFrom
    Pass    = $mailPass
    To      = $mailTo
    Subject = $mailSbj
    Body    = $mailMsg
    ScriptPath = $scriptPath
    Attachment = $log
}
Emailv2 @emailParam