###############################################################################
# SCRIPT NAME: QU-Remove-NonessentialMessages
#
# AUTHOR:
#   John Buxton
#
# PURPOSE:
#   Scans the SQL quarantine table for messages that should not be processed
#   further (SMI='DFD' or containing 'DBLINK1') and removes them in a single
#   batch operation.
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
#
# CHANGE LOG:
#   v1.0 - Initial cleanup and formatting to meet standardized requirements.
#   v1.1 - Added logic-explanation flowerboxes throughout the script to help
#          operators understand what each block is doing and why.
###############################################################################

$version = "v1.0";

# ---------------------------------------------------------
# Open SQL connection for scanning and deletion
# ---------------------------------------------------------
$sqlConn = New-SqlConnection

Write-Host "Scanning for messages with SMI='DFD' or containing 'DBLINK1'...$version" -ForegroundColor Green


# ---------------------------------------------------------
# Count how many messages match the removal criteria
# ---------------------------------------------------------
$countCmd = $sqlConn.CreateCommand()
$countCmd.CommandText = @"
SELECT COUNT(*)
FROM dbo.MsmqMessageQuarantine
WHERE SMI IN ('DFD', '*D*')
   OR BodyText LIKE '%DBLINK1%'
"@

$count = $countCmd.ExecuteScalar()


# ---------------------------------------------------------
# If none found, exit early
# ---------------------------------------------------------
if ($count -eq 0) {
  Write-Host "No 'DFD' or 'DBLINK1' messages found." -ForegroundColor Yellow
}

if ($count -gt 0) {
  Write-Host "Removing $count messages with 'DFD' or 'DBLINK1'..." -ForegroundColor Yellow


  # ---------------------------------------------------------
  # Delete all matching messages in a single SQL operation
  # ---------------------------------------------------------
$delCmd = $sqlConn.CreateCommand()
$delCmd.CommandText = @"
DELETE FROM dbo.MsmqMessageQuarantine
WHERE SMI IN ('DFD', '*D*')
   OR BodyText LIKE '%DBLINK1%';
"@

  $delCmd.ExecuteNonQuery() | Out-Null

  # ---------------------------------------------------------
  # Notify operator of completion
  # ---------------------------------------------------------
  Write-Host "Completed removal of $count messages with 'DFD' and 'DBLINK1'." -ForegroundColor Cyan
}

# ---------------------------------------------------------
# Update all rows where Status='Parsed' → set to 'Ready'
# This prepares messages for the next stage of processing.
# ---------------------------------------------------------
$cmd = $sqlConn.CreateCommand()
$cmd.CommandText = @"
UPDATE dbo.MsmqMessageQuarantine
SET Status = 'Ready'
WHERE Status = 'Parsed';
"@

$cmd.ExecuteNonQuery() | Out-Null

Write-Host "All 'Parsed' messages have been updated to 'Ready'." -ForegroundColor Cyan


# ---------------------------------------------------------
# Close SQL connection
# ---------------------------------------------------------
$sqlConn.Close()


###############################################################################
# END OF SCRIPT
###############################################################################
