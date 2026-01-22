###############################################################################
# SCRIPT NAME: QU-Application
#
# AUTHOR:
#   John Buxton
#
# PURPOSE:
#   Provides an operator-driven interface for processing MSMQ messages through
#   the quarantine pipeline, including database staging, parsing, cleanup,
#   file output, and status review.
#
# PARAMETER (INPUT):
#   None
#
# PARAMETER (OUTPUT):
#   None
#
# RETURN:
#   None
#
# HELPER FUNCTIONS:
#   New-SqlConnection   - Opens and returns an active SQL connection using
#                         Windows Integrated Authentication.
#   Get-QueueDepth      - Returns the current message count in the MSMQ queue.
#   Get-RemainingCount  - Returns the number of rows remaining in the
#                         quarantine table.
#   Show-QueueMonitor   - Displays queue depth samples and returns to menu.
#
# CHANGE LOG:
#   v1.0 (2026-01-10) - Initial release by jBuxton.
#   v1.1 (2026-01-10) - Reformatted header to flower-box style; added helper
#                       function descriptions; removed duplicate logic; cleaned
#                       comments; enforced 2-space indentation.
#   v1.2 (2026-01-11) - Added User Guide menu option and script call.
#   v1.3 (2026-01-13) - Added Queue Count Monitor (Option 6).
#   v1.4 (2026-01-13) - Simplified Queue Monitor to 10 samples with 2-second
#                       interval.
#   v1.5 (2026-01-20) - Removed SQL username/password handling; converted to
#                       Windows Integrated Authentication; removed unused
#                       credential function and globals; cleaned SQL helpers.
###############################################################################


###############################################################################
# MSMQ → SQL → Review → Requeue/Delete Pipeline
# Queue: .\private$\fromnetworkq
###############################################################################

$version = "v1.5"

###############################################################################
# CONFIGURATION
###############################################################################
$QueuePath = ".\private$\fromnetworkq"

[string] $SqlServer = [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName
[string] $Database  = "AircomSrv_Util"

# Load MSMQ assembly
[Reflection.Assembly]::LoadWithPartialName('System.Messaging') | Out-Null


###############################################################################
# FUNCTION NAME: New-SqlConnection
# PURPOSE      : Opens a SQL connection using Windows Integrated Authentication.
###############################################################################
function New-SqlConnection {

  $connString = "Server=$SqlServer;Database=$Database;Integrated Security=True;"
  $conn = New-Object System.Data.SqlClient.SqlConnection $connString
  $conn.Open()
  return $conn
}


###############################################################################
# FUNCTION NAME: Invoke-QuSqlQuery
# PURPOSE      : Executes a SQL query using Windows Integrated Authentication.
# INPUT        :
#   -ServerInstance : SQL Server name or instance
#   -Database       : Target database name
#   -Query          : SQL query text
# OUTPUT       : SQL result set (if any)
###############################################################################
function Invoke-QuSqlQuery {

  param(
    [Parameter(Mandatory = $true)]
    [string]$ServerInstance,

    [Parameter(Mandatory = $true)]
    [string]$Database,

    [Parameter(Mandatory = $true)]
    [string]$Query
  )

  Write-Host "Executing SQL query using Windows Integrated Authentication..." -ForegroundColor Cyan

  try {
    $result = Invoke-Sqlcmd `
      -ServerInstance $ServerInstance `
      -Database $Database `
      -Query $Query

    Write-Host "SQL query completed successfully." -ForegroundColor Green
    return $result
  }
  catch {
    Write-Host "SQL query failed:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
  }
}


###############################################################################
# FUNCTION NAME: Get-QueueDepth
###############################################################################
function Get-QueueDepth {
  $q = New-Object System.Messaging.MessageQueue $QueuePath
  return $q.GetAllMessages().Count
}


###############################################################################
# FUNCTION NAME: Get-RemainingCount
# PURPOSE      : Returns remaining rows in dbo.MsmqMessageQuarantine.
###############################################################################
function Get-RemainingCount {

  $sqlConn = New-SqlConnection
  $cmd = $sqlConn.CreateCommand()
  $cmd.CommandText = "SELECT COUNT(*) FROM dbo.MsmqMessageQuarantine;"
  $count = $cmd.ExecuteScalar()
  $sqlConn.Close()
  return $count
}


###############################################################################
# FUNCTION NAME: Show-QueueMonitor
# PURPOSE:
#   Displays queue depth 5 times, with 2 seconds between each display,
#   then returns automatically to the main menu.
###############################################################################
function Show-QueueMonitor {

  Clear-Host
  Write-Host "==================== QUEUE COUNT MONITOR ====================" -ForegroundColor Cyan
  Write-Host "Displaying queue depth 5 times (every 2 seconds)..." -ForegroundColor Yellow
  Write-Host ""

  for ($i = 1; $i -le 5; $i++) {

    $queueDepth = Get-QueueDepth
    Write-Host ("[{0:00}/5] Queue Depth: {1}" -f $i, $queueDepth) -ForegroundColor Green

    if ($i -lt 5) {
      Start-Sleep -Seconds 2
    }
  }

  Write-Host ""
  Write-Host "Queue monitor complete. Returning to main menu..." -ForegroundColor Cyan
  Start-Sleep -Seconds 2
}


###############################################################################
# MAIN APPLICATION FUNCTION
###############################################################################
function Qu-Application {

  while ($true) {

    Clear-Host
    Write-Host "==================== MSMQ QUARANTINE PROCESSING MENU $version ====================" -ForegroundColor White
    Write-Host "" -ForegroundColor White

    Write-Host -NoNewline "1. " -ForegroundColor Cyan; Write-Host "Move QU Messages Into Quarantine"           -ForegroundColor White
    Write-Host -NoNewline "2. " -ForegroundColor Cyan; Write-Host "Parse DSP DayTime SMI"                      -ForegroundColor White
    Write-Host -NoNewline "3. " -ForegroundColor Cyan; Write-Host "Remove Nonessential Messages (SMI='DFD')"   -ForegroundColor White
    Write-Host -NoNewline "4. " -ForegroundColor Cyan; Write-Host "File QU Messages to Aircom Input Folder"    -ForegroundColor White
    Write-Host -NoNewline "5. " -ForegroundColor Cyan; Write-Host "Display Message Status"                     -ForegroundColor White
    Write-Host -NoNewline "6. " -ForegroundColor Cyan; Write-Host "Display Queue Count"                        -ForegroundColor White
    Write-Host -NoNewline "7. " -ForegroundColor Cyan; Write-Host "Display Services"                           -ForegroundColor White
    Write-Host -NoNewline "8. " -ForegroundColor Cyan; Write-Host "Restart Message Processor Service"          -ForegroundColor White
    Write-Host -NoNewline "U. " -ForegroundColor Cyan; Write-Host "User Guide"                                 -ForegroundColor White
    Write-Host -NoNewline "X. " -ForegroundColor Cyan; Write-Host "Exit"                                       -ForegroundColor White

    Write-Host ""
    Write-Host "===================================================================================" -ForegroundColor White
    Write-Host ""

    $choice = Read-Host "Select an option"

    switch ($choice.ToUpper()) {

      "1" {
        D:\PS_Scripts\QU-Move-AllMsmqMessagesToDatabase.ps1
        Pause
      }

      "2" {
        D:\PS_Scripts\QU-Parse-MsmqFields.ps1
        Pause
      }

      "3" {
        D:\PS_Scripts\QU-Remove-NonessentialMessages.ps1
        Pause
      }

      "4" {
        D:\PS_Scripts\QU-File-AllMessagesAdaptive.ps1
        Pause
      }

      "5" {
        D:\PS_Scripts\QU-Display-MessageStatusSummary.ps1
        Pause
      }

      "6" {
        Show-QueueMonitor
        Pause
      }

      "7" {
        $myString = Get-Service |
          Select-Object StartType, Status, Name, DisplayName |
          Where-Object { $_.DisplayName -like "AS*" } |
          Out-String

        $myString = $myString.Substring(2, $myString.Length - 6)
        Write-Host $myString

        $myString = Get-Service |
          Select-Object StartType, Status, Name, DisplayName |
          Where-Object { $_.DisplayName -like "*AIRCOM_DB*" } |
          Out-String

        Write-Host $myString
        Pause
      }

      "8" {
        . "D:\PS_Scripts\QU-Restart-AsMessageProcessor.ps1"
        Restart-AsMessageProcessor -Force N
        Pause
      }

      "U" {
        D:\PS_Scripts\QU-UserGuide.ps1
        Pause
      }

      "X" {
        Write-Host "Exiting..." -ForegroundColor Green
        return
      }

      default {
        Write-Host "Invalid selection. Try again." -ForegroundColor Red
        Start-Sleep -Seconds 1
      }
    }
  }
}

###############################################################################
# SCRIPT ENTRY POINT
###############################################################################
Clear-Host
Qu-Application

###############################################################################
# END OF SCRIPT
###############################################################################
