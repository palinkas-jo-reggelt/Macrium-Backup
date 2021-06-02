<#

.SYNOPSIS
	Macrium Reflect Backup

.DESCRIPTION
	Script to run and send notifications for Macrium Reflect

.FUNCTIONALITY
	* Runs Macrium Reflect using specified XML config
	* Copies Macrium html log file to webserver
	* Creates short link to log file
	* Emails report with link to log file

.PARAMETER 
	-s 
	-full = full backup
	-inc  = incremental backup
	-diff = differential backup
	
.NOTES
	* Run daily from task scheduler with administrator privileges 
	* Requires: Macrium Reflect 7+
	* Requires: web server
	* Requires: YoURLs short link creator
	
.EXAMPLE
	PS C:\Windows\system32> C:\scripts\Reflect\DailyBackup.ps1 -diff

#>

<###   SCRIPT PARAMETERS   ###>
Param([switch]$s, [switch]$full, [switch]$inc, [switch]$diff)

<###   USER VARIABLES   ###>
$strReflectPath    = "C:\Program Files\Macrium\Reflect\reflect.exe"
$strXmlFilePath    = "C:\Users\user\Documents\Reflect\DailyBackup.xml"
$EmailFrom         = "notify@mydomain.tld"
$EmailTo           = "admin@mydomain.tld"
$Subject           = "Macrium Nightly Backup"
$SMTPServer        = "mydomain.tld"
$SMTPAuthUser      = "notify@mydomain.tld"
$SMTPAuthPass      = "supersecretpassword"
$SMTPPort          =  587
$SSL               = $True                                  # If true, will use tls connection to send email
$UseHTML           = $True                                  # If true, will format and send email body as html (with color!)
$AttachDebugLog    = $True                                  # If true, will attach debug log to email report - must also select $VerboseFile
$MaxAttachmentSize = 1                                      # Size in MB
$MacriumLogDir     = "C:\ProgramData\Macrium\Reflect"       # Default location - change if necessary
$WebDir            = "C:\xampp\htdocs\mydomain.tld\maclog"  # File location of web dir to place log files for online viewing
$WebBaseURL        = "https://maclog.mydomain.tld/"         # Please leave trailing slash "/"
$YoURLsToken       = 'secrettoken'                          # YoURLs API token
$YoURLsURI         = 'https://url.mydomain.tld'             # YoURLs Base URL

<###   FUNCTIONS   ###>
Function Email ($EmailOutput) {
	If ($UseHTML){
		If ($EmailOutput -match "\[OK\]") {$EmailOutput = $EmailOutput -Replace "\[OK\]","<span style=`"background-color:green;color:white;font-weight:bold;font-family:Courier New;`">[OK]</span>"}
		If ($EmailOutput -match "\[INFO\]") {$EmailOutput = $EmailOutput -Replace "\[INFO\]","<span style=`"background-color:yellow;font-weight:bold;font-family:Courier New;`">[INFO]</span>"}
		If ($EmailOutput -match "\[ERROR\]") {$EmailOutput = $EmailOutput -Replace "\[ERROR\]","<span style=`"background-color:red;color:white;font-weight:bold;font-family:Courier New;`">[ERROR]</span>"}
		If ($EmailOutput -match "^\s$") {$EmailOutput = $EmailOutput -Replace "\s","&nbsp;"}
		Write-Output "<tr><td>$EmailOutput</td></tr>" | Out-File $EmailBody -Encoding ASCII -Append
	} Else {
		Write-Output $EmailOutput | Out-File $EmailBody -Encoding ASCII -Append
	}	
}

Function EmailResults {
	Try {
		$Body = (Get-Content -Path $EmailBody | Out-String )
		$Message = New-Object System.Net.Mail.Mailmessage $EmailFrom, $EmailTo, $Subject, $Body
		$Message.IsBodyHTML = $UseHTML
		$SMTP = New-Object System.Net.Mail.SMTPClient $SMTPServer,$SMTPPort
		$SMTP.EnableSsl = $SSL
		$SMTP.Credentials = New-Object System.Net.NetworkCredential($SMTPAuthUser, $SMTPAuthPass); 
		$SMTP.Send($Message)
	}
	Catch {
		Write-Output "$(Get-Date -f G) : [ERROR] EmailResults Function : $($Error[0])" | Out-File "$PSScriptRoot\MacriumErrorLog-$DateString.log" -Encoding ASCII -Append
	}
}

Function ElapsedTime ($EndTime) {
	$TimeSpan = New-Timespan $EndTime
	If (([int]($TimeSpan).Hours) -eq 0) {$Hours = ""} ElseIf (([int]($TimeSpan).Hours) -eq 1) {$Hours = "1 hour "} Else {$Hours = "$([int]($TimeSpan).Hours) hours "}
	If (([int]($TimeSpan).Minutes) -eq 0) {$Minutes = ""} ElseIf (([int]($TimeSpan).Minutes) -eq 1) {$Minutes = "1 minute "} Else {$Minutes = "$([int]($TimeSpan).Minutes) minutes "}
	If (([int]($TimeSpan).Seconds) -eq 1) {$Seconds = "1 second"} Else {$Seconds = "$([int]($TimeSpan).Seconds) seconds"}
	
	If (($TimeSpan).TotalSeconds -lt 1) {
		$Return = "less than 1 second"
	} Else {
		$Return = "$Hours$Minutes$Seconds"
	}
	Return $Return
}

Function CopyLogFile {
	$MacriumLog = (Get-ChildItem $MacriumLogDir -Attributes !Directory *.html | Sort-Object -Descending -Property LastWriteTime | Select -First 1)
	Copy-Item $MacriumLog.FullName -Destination $WebDir
	Add-Type -AssemblyName System.Web
	$LongURL = [System.Web.HTTPUtility]::UrlEncode($WebBaseURL + $MacriumLog.Name)
	$YoURLsAPI = $YoURLsURI + "/yourls-api.php?title=MacriumLog&signature=" + $YoURLsToken + "&action=shorturl&format=json&url=" + $LongURL
	$ShortURL = (Invoke-RestMethod $YoURLsAPI).shorturl
	Return $ShortURL
}

Function Backup(){
	$strType = GetBackupTypeParameter
	$strArgs = "-e -w $strType `"$strXmlFilePath`""
	$RunBackup = Start-Process -FilePath $strReflectPath -ArgumentList $strArgs -PassThru -Wait
	$iResult = $RunBackup.ExitCode
	switch ($iResult){
		2 { Email "[ERROR] XML invalid - See log for error";   break; }
		1 { Email "[ERROR] Backup failed - See log for error"; break; }
		0 { Email "[OK] Backup succeeded";                     break; }
	}
	return $iResult
}

Function GetBackupTypeParameter(){
	if ($full -eq $true) { return '-full'; }
	if ($inc  -eq $true) { return '-inc';  }
	if ($diff -eq $true) { return '-diff'; }
	return ''; # Clone
}


<###   START SCRIPT   ###>

$StartScript = Get-Date
$DateString = (Get-Date).ToString("yyyy-MM-dd")

<#  Delete old debug files and create new  #>
$EmailBody = "$PSScriptRoot\EmailBody.log"
If (Test-Path $EmailBody) {Remove-Item -Force -Path $EmailBody}
New-Item $EmailBody
If ($UseHTML) {
	Write-Output "
		<!DOCTYPE html><html>
		<head><meta name=`"viewport`" content=`"width=device-width, initial-scale=1.0 `" /></head>
		<body style=`"font-family:Arial Narrow`"><table>
	" | Out-File $EmailBody -Encoding ASCII -Append
}
If ($UseHTML) {
	Email "<center>:::&nbsp;&nbsp;&nbsp;Macrium Backup Routine&nbsp;&nbsp;&nbsp;:::</center>"
	Email "<center>$(Get-Date -f D)</center>"
	Email " "
} Else {
	Email ":::   Macrium Backup Routine   :::"
	Email "       $(Get-Date -f D)"
	Email " "
	Email " "
}

<#  Execute backup  #>
Email "Macrium Reflect Backup Definition File:"
Email "$strXmlFilePath"
Email " "
$iExitCode = Backup
Email " "
Email "Script finished with exit code $iExitCode."

<#  Finish up and send email  #>
Email "Macrium Backup routine completed in $(ElapsedTime $StartScript)"
Email " "
$LogURILong = CopyLogFile
$LogURIShort = $LogURILong -Replace "https://", ""
If ($UseHTML) { Email "Debug Log: <a href=`"$LogURILong`">$LogURIShort</a>" } Else { Email "Debug Log: $LogURILong" }
If ($UseHTML) {Write-Output "</table></body></html>" | Out-File $EmailBody -Encoding ASCII -Append}
EmailResults

<#  Exit with exit code  #>
Exit $iExitCode
