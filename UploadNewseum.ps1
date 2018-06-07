<#

UploadNewseum.ps1

    2018-06-05 Initial Creation

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



#$ftp    = Get-WJFTP -Name Newseum | Where-Object{$_.User -eq 'ny_wjny'}
$ftp    = Get-WJFTP -Name WorldJournalNewYork
$ePaper = Get-WJPath -Name epaper
$workDate = ((Get-Date).AddDays(0)).ToString("yyyyMMdd")
Write-Log -Verb "ftp" -Noun $ftp.Path -Path $log -Type Short -Status Normal
Write-Log -Verb "ePaper" -Noun $ePaper.Path -Path $log -Type Short -Status Normal
Write-Log -Verb "workDate" -Noun $workDate -Path $log -Type Short -Status Normal


$localFileName  = "NY" + $workDate + "A01.pdf"
$localFilePath  = $ePaper.Path + $workDate + "\optimizeda\" + $localFileName
$localFilePath2 = "C:\temp\" + $localFileName
$remoteFileName = "NY_WJNY_" + $workDate + ".pdf"
$remoteFilePath = $ftp.Path + $remoteFileName

Write-Log -Verb "localFileName" -Noun $localFileName -Path $log -Type Short -Status Normal
Write-Log -Verb "localFilePath" -Noun $localFilePath -Path $log -Type Short -Status Normal
Write-Log -Verb "remoteFileName" -Noun $remoteFileName -Path $log -Type Short -Status Normal
Write-Log -Verb "remoteFilePath" -Noun $remoteFilePath -Path $log -Type Short -Status Normal

$upload = WebClient-UploadFile -Username $ftp.User -Password $ftp.Pass -RemoteFilePath $remoteFilePath -LocalFilePath $localFilePath
Write-Log -Verb $upload.Verb -Noun $upload.Noun -Path $log -Type Long -Status $upload.Status

$download = WebClient-DownloadFile -Username $ftp.User -Password $ftp.Pass -RemoteFilePath $remoteFilePath -LocalFilePath $localFilePath2
Write-Log -Verb $download.Verb -Noun $download.Noun -Path $log -Type Long -Status $download.Status



###################################################################################

Write-Line -Length 100 -Path $log
Write-Log -Verb "LOG END" -Noun $log -Path $log -Type Long -Status Normal
if($hasError){ $mailSbj = "ERROR " + $scriptName }

$emailParam = @{
    From    = $mailFrom
    Pass    = $mailPass
    To      = $mailTo
    Subject = $mailSbj
    Body    = $scriptName + " completed at: " + (Get-Date).ToString("HH:mm:ss") + "`n`n" + $mailMsg
    ScriptPath = $scriptPath
    Attachment = $log
}
#Emailv2 @emailParam
$emailParam.Body