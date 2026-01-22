###############################################################################
# FUNCTION NAME: QU-Restart-AsMessageProcessor
#
# PURPOSE:
#   Provides a safe, operator-friendly workflow for restarting the
#   "AS Message Processor" service. Supports optional -Force parameter:
#     • If -Force Y → Restart immediately with no operator prompt
#     • If -Force not supplied → Ask operator for confirmation
#
# PARAMETERS:
#   [string] $Force  - Optional. If "Y", restart occurs without prompting.
#
# USAGE:
#   Restart-AsMessageProcessor
#   Restart-AsMessageProcessor -Force Y
#
# CHANGE LOG:
#   v1.0 - Initial standalone version extracted from QU-Application.
#   v1.1 - Added improved error handling and status reporting.
#   v1.2 - Added operator prompt and consistent color formatting.
#   v1.3 - Added optional -Force parameter for non-interactive restart.
###############################################################################

function Restart-AsMessageProcessor {
  param(
    [string] $Force = ""
  )

  $serviceName = "AS Message Processor"

  Write-Host ""
  Write-Host "Restart service for '$serviceName'" -ForegroundColor Green
  Write-Host "------------------------------------------------------------"

  # ---------------------------------------------------------
  # Query service status
  # ---------------------------------------------------------
  try {
    $svc = Get-Service -DisplayName $serviceName -ErrorAction Stop
    Write-Host "Current Status: $($svc.Status)" -ForegroundColor Yellow
  }
  catch {
    Write-Host "ERROR: Unable to query service '$serviceName'." -ForegroundColor Red
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor DarkRed
    return
  }

  # ---------------------------------------------------------
  # Determine whether to prompt operator
  # ---------------------------------------------------------
  $autoRestart = ($Force.ToUpper() -eq "Y")

  if (-not $autoRestart) {
    $choice = Read-Host "Do you want to restart this service? (Y/N)"
    if ($choice.ToUpper() -ne "Y") {
      Write-Host "Service restart canceled by operator." -ForegroundColor Cyan
      return
    }
  }
  else {
    Write-Host "Force parameter detected — restarting without operator prompt." -ForegroundColor Cyan
  }

  # ---------------------------------------------------------
  # Refresh status before acting
  # ---------------------------------------------------------
  $svc = Get-Service -DisplayName $serviceName

  # ---------------------------------------------------------
  # If STOPPED → START
  # ---------------------------------------------------------
  if ($svc.Status -eq 'Stopped') {

    Write-Host "Service is currently Stopped. Starting service..." -ForegroundColor Yellow

    try {
      Start-Service -DisplayName $serviceName -ErrorAction Stop
      Write-Host "Service '$serviceName' started successfully." -ForegroundColor Green
    }
    catch {
      Write-Host "ERROR: Failed to start service '$serviceName'." -ForegroundColor Red
      Write-Host "Details: $($_.Exception.Message)" -ForegroundColor DarkRed
    }

    return
  }

  # ---------------------------------------------------------
  # If RUNNING → STOP then START
  # ---------------------------------------------------------
  if ($svc.Status -eq 'Running') {

    Write-Host "Stopping service '$serviceName'..." -ForegroundColor Yellow

    try {
      Stop-Service -DisplayName $serviceName -Force -ErrorAction Stop
      Write-Host "Service '$serviceName' stopped successfully." -ForegroundColor Green
    }
    catch {
      Write-Host "ERROR: Failed to stop service '$serviceName'." -ForegroundColor Red
      Write-Host "Details: $($_.Exception.Message)" -ForegroundColor DarkRed

      # -----------------------------------------------------
      # Fallback: Attempt START anyway
      # -----------------------------------------------------
      Write-Host "Attempting to start service anyway..." -ForegroundColor Yellow
      try {
        Start-Service -DisplayName $serviceName -ErrorAction Stop
        Write-Host "Service '$serviceName' started successfully." -ForegroundColor Green
      }
      catch {
        Write-Host "ERROR: Failed to start service '$serviceName'." -ForegroundColor Red
        Write-Host "Details: $($_.Exception.Message)" -ForegroundColor DarkRed
      }

      return
    }

    # ---------------------------------------------------------
    # Start after successful stop
    # ---------------------------------------------------------
    Write-Host "Starting service '$serviceName'..." -ForegroundColor Yellow

    try {
      Start-Service -DisplayName $serviceName -ErrorAction Stop
      Write-Host "Service '$serviceName' started successfully." -ForegroundColor Green
    }
    catch {
      Write-Host "ERROR: Failed to start service '$serviceName' after stop." -ForegroundColor Red
      Write-Host "Details: $($_.Exception.Message)" -ForegroundColor DarkRed
    }

    return
  }

  # ---------------------------------------------------------
  # Any other state → warn operator
  # ---------------------------------------------------------
  Write-Host "Service is in unexpected state '$($svc.Status)'." -ForegroundColor Red
  Write-Host "Manual intervention may be required." -ForegroundColor DarkRed
}

###############################################################################
# END OF SCRIPT
###############################################################################
