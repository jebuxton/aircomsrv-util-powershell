###############################################################################
# SCRIPT NAME: QU-Move-AllMsmqMessagesToDatabase
#
# AUTHOR:
#   John Buxton
#
# PURPOSE:
#   Reads ALL messages currently on the MSMQ queue and writes them into the
#   SQL quarantine table. No batching, no operator prompts. If any messages
#   were processed, the script force-restarts the AS Message Processor service.
#
# CHANGE LOG:
#   v1.0 - Initial formatting and cleanup.
#   v1.1 - Added logic-explanation flowerboxes.
#   v1.2 - Added batch-loop logic with operator prompts.
#   v1.3 - Added $messagesWereProcessed flag.
#   v2.0 - Removed batch loop and operator prompts. Script now processes all
#          messages in a single pass and force-restarts the service if needed.
###############################################################################

$version = "v2.0";

###############################################################################
# FUNCTION: Save-MsmqMessageToDatabase
###############################################################################
function Save-MsmqMessageToDatabase {
  param(
    [System.Data.SqlClient.SqlConnection] $SqlConnection,
    [string] $BodyText,
    [byte[]] $BodyBinary
  )

  if ($BodyText -like "*DBLINK1*") {
    return
  }

  $cmd = $SqlConnection.CreateCommand()

  $cmd.CommandText = @"
INSERT INTO dbo.MsmqMessageQuarantine
(QueuePath, BodyText, BodyBinary, Status)
VALUES (@QueuePath, @BodyText, @BodyBinary, 'Pending');
"@

  $cmd.Parameters.AddWithValue("@QueuePath", $QueuePath) | Out-Null

  if ($BodyText) {
    $cmd.Parameters.AddWithValue("@BodyText", $BodyText) | Out-Null
  }
  else {
    $cmd.Parameters.AddWithValue("@BodyText", [DBNull]::Value) | Out-Null
  }

  $paramPayload = $cmd.Parameters.Add("@BodyBinary", [System.Data.SqlDbType]::VarBinary, -1)
  $paramPayload.Value = $BodyBinary

  $cmd.ExecuteNonQuery() | Out-Null
}

###############################################################################
# STEP 1 — PROCESS ALL MSMQ MESSAGES IN ONE PASS
###############################################################################

Write-Host "Opening MSMQ queue: $QueuePath $version" -ForegroundColor Green

# Recreate queue object for accurate count
if ($queue) { $queue.Close() }
$queue = New-Object System.Messaging.MessageQueue $QueuePath
$queue.Formatter = New-Object System.Messaging.BinaryMessageFormatter

# Count all messages currently on the queue
$msgCount = $queue.GetAllMessages().Count
Write-Host "Messages currently on queue: $msgCount" -ForegroundColor Yellow

if ($msgCount -eq 0) {
  Write-Host "No messages to move." -ForegroundColor Cyan
  Write-Host "Service restart skipped." -ForegroundColor Cyan
  return
}

# Open SQL connection
$sqlConn = New-SqlConnection
$messagesProcessed = 0

# Process exactly $msgCount messages
for ($i = 1; $i -le $msgCount; $i++) {

  try {
    $msg = $queue.Receive([TimeSpan]::FromSeconds(2))
  }
  catch {
    Write-Host "Receive() threw: $($_.Exception.Message)" -ForegroundColor DarkRed
    break
  }

  if ($msg -eq $null) { break }

  $stream = New-Object System.IO.MemoryStream
  $msg.BodyStream.CopyTo($stream)
  $bodyBinary = $stream.ToArray()

  $bodyText = [System.Text.Encoding]::ASCII.GetString($bodyBinary)

  Save-MsmqMessageToDatabase `
    -SqlConnection $sqlConn `
    -BodyText $bodyText `
    -BodyBinary $bodyBinary

  $messagesProcessed++
}

$sqlConn.Close()

Write-Host "Total messages moved: $messagesProcessed" -ForegroundColor Cyan

###############################################################################
# STEP 2 — CONDITIONAL SERVICE RESTART
###############################################################################

if ($messagesProcessed -gt 0) {

  Write-Host "Restarting AS Message Processor service..." -ForegroundColor Yellow

  # Dot-source the restart function
  . "D:\PS_Scripts\QU-Restart-AsMessageProcessor.ps1"

  # Force restart (no operator prompt)
  Restart-AsMessageProcessor -Force Y
}
else {
  Write-Host "No messages were processed — service restart skipped." -ForegroundColor Cyan
}

###############################################################################
# END OF SCRIPT
###############################################################################
