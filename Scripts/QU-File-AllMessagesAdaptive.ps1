###############################################################################
# SCRIPT NAME: QU-File-AllMessagesAdaptive
#
# AUTHOR:
#   John Buxton
#
# PURPOSE:
#   Reads Parsed messages from the SQL quarantine table, writes each message
#   to a file in the Aircom input directory, deletes the SQL row, and
#   dynamically adjusts batch size based on MSMQ queue depth to prevent
#   downstream overload. Now includes full reporting metrics.
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
#   New-SqlConnection - Opens and returns an active SQL connection.
#   Get-QueueDepth    - Returns the current MSMQ queue depth.
#
# CHANGE LOG:
#   v1.0 - Initial cleanup and formatting.
#   v1.1 - Added logic-explanation flowerboxes.
#   v1.2 - Added batch file count + remaining message count to queue-depth output.
#   v1.3 - Added full reporting metrics (runtime, averages, safety stops, etc.)
#   v1.4 - Added colorized final report (Option A: simple & clean)
###############################################################################

$version = "v1.4";

###############################################################################
# INITIALIZATION
###############################################################################

$sqlConn = New-SqlConnection

$outDir = "D:\input"
if (-not (Test-Path $outDir)) {
  New-Item -ItemType Directory -Path $outDir | Out-Null
}

###############################################################################
# USER CONFIRMATION — PROCEED WITH MOVING FILES TO AIRCOM?
###############################################################################

Write-Host ""
Write-Host "This operation will move all READY messages to the Aircom input folder:" -ForegroundColor Yellow
Write-Host "    $outDir" -ForegroundColor Cyan
Write-Host ""
$confirm = Read-Host "Do you want to proceed? (Y/N)"

if ($confirm.ToUpper() -ne "Y") {
  Write-Host "Operation canceled by user. No messages were moved." -ForegroundColor Cyan
  return
}

Write-Host "Proceeding with Aircom file output..." -ForegroundColor Green
Write-Host ""



$batchSize = 20
$global:RequeueSeq = 0

Write-Host "Adaptive file-output engine starting...$version" -ForegroundColor Green

###############################################################################
# REPORTING METRICS INITIALIZATION
###############################################################################

$startTime = Get-Date
$totalMessages = 0
$totalBatches = 0
$safetyStops = 0
$maxQueueDepth = 0
$totalSafetyWaitSeconds = 0
$maxBatchSize = $batchSize

###############################################################################
# MAIN PROCESSING LOOP
###############################################################################
while ($true) {

  # ---------------------------------------------------------
  # Reset per-batch file counter
  # ---------------------------------------------------------
  $filesWrittenThisBatch = 0
  $totalBatches++

  # ---------------------------------------------------------
  # Pull next batch of Ready messages from SQL
  # ---------------------------------------------------------
  $cmd = $sqlConn.CreateCommand()
  $cmd.CommandText = @"
SELECT TOP($batchSize) Id, BodyText, SMI, ParsedTimestamp
FROM dbo.MsmqMessageQuarantine
WHERE Status = 'Ready'
ORDER BY Id;
"@

  $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
  $dt = New-Object System.Data.DataTable
  $adapter.Fill($dt) | Out-Null

  if ($dt.Rows.Count -eq 0) {
    Write-Host "All messages have been processed." -ForegroundColor Green
    break
  }

  #############################################################################
  # PROCESS EACH MESSAGE IN THE CURRENT BATCH
  #############################################################################
  foreach ($row in $dt.Rows) {

    $SMI = $row.SMI
    $ts  = $row.ParsedTimestamp

    $dayHHMM = "{0:00}{1:00}{2:00}" -f $ts.Day, $ts.Hour, $ts.Minute

    $seq = $global:RequeueSeq
    $global:RequeueSeq++
    if ($global:RequeueSeq -gt 99999) { $global:RequeueSeq = 0 }

    $fileName = "Requeue_{0}_{1}_{2:00000}.txt" -f $SMI, $dayHHMM, $seq
    $filePath = Join-Path $outDir $fileName

    try {
      Set-Content -Path $filePath -Value $row.BodyText
      $filesWrittenThisBatch++
      $totalMessages++
    }
    catch {
      Write-Host "ERROR writing file for ID $($row.Id): $_" -ForegroundColor Red
      continue
    }

    # Delete SQL row
    $upd = $sqlConn.CreateCommand()
    $upd.CommandText = @"
DELETE FROM dbo.MsmqMessageQuarantine
WHERE Id = @Id;
"@
    $upd.Parameters.AddWithValue("@Id", $row.Id) | Out-Null
    $upd.ExecuteNonQuery() | Out-Null
  }

  #############################################################################
  # THROTTLING + SAFETY CONTROLS
  #############################################################################

  Start-Sleep -Seconds 6

  # ---------------------------------------------------------
  # Count remaining messages in SQL (Status='Ready')
  # ---------------------------------------------------------
  $remainCmd = $sqlConn.CreateCommand()
  $remainCmd.CommandText = "SELECT COUNT(*) FROM dbo.MsmqMessageQuarantine WHERE Status='Ready';"
  $remainingMessages = $remainCmd.ExecuteScalar()

  # ---------------------------------------------------------
  # Check MSMQ queue depth and display all metrics
  # ---------------------------------------------------------
  $queueDepth = Get-QueueDepth

  if ($queueDepth -gt $maxQueueDepth) { $maxQueueDepth = $queueDepth }

  Write-Host "Queue depth check: $queueDepth | Files written this batch: $filesWrittenThisBatch | Remaining messages: $remainingMessages" -ForegroundColor Cyan

  # Warning threshold
  if ($queueDepth -ge 150 -and $queueDepth -le 200) {
    Write-Host "WARNING: Queue depth is $queueDepth (≥150). Approaching safety limit." -ForegroundColor Yellow
  }

  # Hard stop threshold
  if ($queueDepth -gt 200) {
    $safetyStops++
    Write-Host "Queue depth exceeded 200 → STOPPING processing for safety." -ForegroundColor Red
    Write-Host "Waiting for queue to drain below 50 before resuming..." -ForegroundColor White

    while ($true) {
      Start-Sleep -Seconds 6
      $totalSafetyWaitSeconds += 6

      $queueDepth = Get-QueueDepth
      if ($queueDepth -gt $maxQueueDepth) { $maxQueueDepth = $queueDepth }

      Write-Host "Queue depth during wait: $queueDepth" -ForegroundColor DarkYellow

      if ($queueDepth -lt 50) {
        Write-Host "Queue depth is now $queueDepth (<50). Resuming processing." -ForegroundColor Green
        break
      }
    }
  }

  #############################################################################
  # ADAPTIVE BATCH SIZE LOGIC
  #############################################################################
  if ($queueDepth -le 100) {
    $batchSize += 10
    if ($batchSize -gt $maxBatchSize) { $maxBatchSize = $batchSize }
    Write-Host "Queue depth ≤ 100 → increasing batch size to $batchSize" -ForegroundColor Yellow
  }
  else {
    $batchSize = 20
    Write-Host "Queue depth > 100 → resetting batch size to 20" -ForegroundColor White
  }
}

###############################################################################
# CLEANUP
###############################################################################
$sqlConn.Close()

###############################################################################
# FINAL REPORT (Colorized — Option A)
###############################################################################

$endTime = Get-Date
$runTime = $endTime - $startTime
$minutes = [math]::Max(1, [int]$runTime.TotalMinutes)

$msgsPerMin   = [math]::Round($totalMessages / $minutes, 2)
$msgsPerBatch = if ($totalBatches -gt 0) { 
    [math]::Round($totalMessages / $totalBatches, 2) 
} else { 
    0 
}

Write-Host ""
Write-Host "==================== Aircom File Input Report $version ===========" -ForegroundColor Cyan

Write-Host -NoNewline "         Start Time:  " -ForegroundColor White
Write-Host ("{0:HH:mm}" -f $startTime) -ForegroundColor Green

Write-Host -NoNewline "           End Time:  " -ForegroundColor White
Write-Host ("{0:HH:mm}" -f $endTime) -ForegroundColor Green

Write-Host -NoNewline "           Run Time:  " -ForegroundColor White
Write-Host ("{0:hh\:mm}" -f $runTime) -ForegroundColor Green

Write-Host -NoNewline "Total Files Written:  " -ForegroundColor White
Write-Host ("{0:N0}" -f $totalMessages) -ForegroundColor Green

Write-Host -NoNewline "       Msgs per min:  " -ForegroundColor White
Write-Host ("{0}" -f $msgsPerMin) -ForegroundColor Green

Write-Host -NoNewline "     Msgs per Batch:  " -ForegroundColor White
Write-Host ("{0}" -f $msgsPerBatch) -ForegroundColor Green

Write-Host -NoNewline "       Safety Stops:  " -ForegroundColor White
Write-Host ("{0}" -f $safetyStops) -ForegroundColor Green

Write-Host -NoNewline "     Max QueueDepth:  " -ForegroundColor White
Write-Host ("{0}" -f $maxQueueDepth) -ForegroundColor Green

Write-Host -NoNewline "        Safety Wait:  " -ForegroundColor White
Write-Host ("{0} sec" -f $totalSafetyWaitSeconds) -ForegroundColor Green

Write-Host -NoNewline "      Max BatchSize:  " -ForegroundColor White
Write-Host ("{0}" -f $maxBatchSize) -ForegroundColor Green

Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host ""


###############################################################################
# END OF SCRIPT
###############################################################################

