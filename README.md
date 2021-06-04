# Macrium Backup
 Macrium Backup - Powershell

### .SYNOPSIS
 Macrium Reflect Daily Backup

### .DESCRIPTION
 Script to run and send notifications for Macrium Reflect

### .FUNCTIONALITY
 * Frees up space on target drive if necessary
 * Runs Macrium Reflect using specified XML config
 * Copies Macrium html log file to webserver
 * Creates short link to log file
 * Emails report with link to log file

### .PARAMETER 
 * -s      = ?
 * -full   = full backup
 * -inc    = incremental backup
 * -diff   = differential backup
 * [empty] = clone
	
### .NOTES
 Run daily from task scheduler with administrator privileges 
	
### .EXAMPLE
 PS C:\Windows\system32> C:\scripts\Reflect\DailyBackup.ps1 -diff
