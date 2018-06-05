<#

UploadNewseum.ps1

    2018-06-05 Initial Creation

#>

if (!($env:PSModulePath -match 'C:\\PowerShell\\_Modules')) {
    $env:PSModulePath = $env:PSModulePath + ';C:\PowerShell\_Modules\'
}

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

$localFileName  = "NY" + $workDate + "A01.pdf"
$localFilePath  = $ePaper.Path + $workDate + "\optimizeda\" + $localFileName
$remoteFileName = "NY_WJNY_" + $workDate + ".pdf"
#$remoteFilePath = $ftp.Path + $remoteFileName
$remoteFilePath = $ftp.Path + "temp\" +  $remoteFileName

$webClient = New-Object System.Net.WebClient 
$webClient.Credentials = New-Object System.Net.NetworkCredential($ftp.User, $ftp.Pass)  
$uri       = New-Object System.Uri($remoteFilePath) 
$webClient.UploadFile($uri, $localFilePath)



###################################################################################

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
Emailv2 @emailParam