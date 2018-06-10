<#

UploadNewseum.ps1

    2018-06-05 Initial Creation
    2018-06-08 Ready for production

#>

if (!($env:PSModulePath -match 'C:\\PowerShell\\_Modules')) {
    $env:PSModulePath = $env:PSModulePath + ';C:\PowerShell\_Modules\'
}

Import-Module WorldJournal.Ftp -Verbose -Force
Import-Module WorldJournal.Log -Verbose -Force
Import-Module WorldJournal.Email -Verbose -Force
Import-Module WorldJournal.Server -Verbose -Force

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

Write-Log -Verb "LOG START" -Noun $log -Path $log -Type Long -Status Normal
Write-Line -Length 100 -Path $log

###################################################################################



#$ftp    = Get-WJFTP -Name WorldJournalNewYork
$ftp    = Get-WJFTP -Name Newseum | Where-Object{$_.User -eq 'ny_wjny'}
$ePaper = Get-WJPath -Name epaper
$workDate = ((Get-Date).AddDays(0)).ToString("yyyyMMdd")
Write-Log -Verb "ftp" -Noun $ftp.Path -Path $log -Type Short -Status Normal
Write-Log -Verb "ePaper" -Noun $ePaper.Path -Path $log -Type Short -Status Normal
Write-Log -Verb "workDate" -Noun $workDate -Path $log -Type Short -Status Normal

$localTemp = "C:\temp\" + $scriptName + "\"
if (!(Test-Path($localTemp))) {New-Item $localTemp -Type Directory | Out-Null}

$localFileName  = "NY" + $workDate + "A01.pdf"
$localFilePath  = $ePaper.Path + $workDate + "\optimizeda\" + $localFileName
$localFilePath2 = $localTemp + $localFileName
$remoteFileName = "NY_WJNY_" + $workDate + ".pdf"
$remoteFilePath = $ftp.Path + $remoteFileName

Write-Log -Verb "localFileName" -Noun $localFileName -Path $log -Type Short -Status Normal
Write-Log -Verb "localFilePath" -Noun $localFilePath -Path $log -Type Short -Status Normal
Write-Log -Verb "remoteFileName" -Noun $remoteFileName -Path $log -Type Short -Status Normal
Write-Log -Verb "remoteFilePath" -Noun $remoteFilePath -Path $log -Type Short -Status Normal

Write-Log -Verb "UPLOAD FROM" -Noun $localFilePath -Path $log -Type Long -Status Normal
Write-Log -Verb "UPLOAD TO" -Noun $remoteFilePath -Path $log -Type Long -Status Normal
$upload = WebClient-UploadFile -Username $ftp.User -Password $ftp.Pass -RemoteFilePath $remoteFilePath -LocalFilePath $localFilePath
$mailMsg = $mailMsg + (Write-Log -Verb $upload.Verb -Noun $upload.Noun -Path $log -Type Long -Status $upload.Status -Output String) + "`n"

Write-Log -Verb "DOWNLOAD FROM" -Noun $remoteFilePath -Path $log -Type Long -Status Normal
Write-Log -Verb "DOWNLOAD TO" -Noun $localFilePath2 -Path $log -Type Long -Status Normal
$download = WebClient-DownloadFile -Username $ftp.User -Password $ftp.Pass -RemoteFilePath $remoteFilePath -LocalFilePath $localFilePath2
$mailMsg = $mailMsg + (Write-Log -Verb $download.Verb -Noun $download.Noun -Path $log -Type Long -Status $download.Status -Output String) + "`n"

Write-Log -Verb "REMOVING" -Noun $localFilePath2 -Path $log -Type Long -Status Normal
try{
    $temp = $localFilePath2
    Remove-Item $localFilePath2 -Force
    $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun "Complete" -Path $log -Type Long -Status Good -Output String)
}catch{
    $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun "Complete" -Path $log -Type Long -Status Bad -Output String)
}

if( ($upload.Status -eq 'Bad') -or ($download.Status -eq 'Bad') ){
    $hasError = $true
}



###################################################################################

Write-Line -Length 100 -Path $log
Write-Log -Verb "LOG END" -Noun $log -Path $log -Type Long -Status Normal
if($hasError){ $mailSbj = "ERROR " + $scriptName }

$emailParam = @{
    From    = $mailFrom
    Pass    = $mailPass
    To      = $mailTo
    Subject = $mailSbj
    Body    = $scriptName + " completed at " + (Get-Date).ToString("HH:mm:ss") + "`n`n" + $mailMsg
    ScriptPath = $scriptPath
    Attachment = $log
}
Emailv2 @emailParam