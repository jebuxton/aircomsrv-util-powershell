###############################################################################
# SCRIPT NAME: QU-Parse-MsmqFields
#
# AUTHOR:
#   John Buxton
#
# PURPOSE:
#   Reads all Pending messages from the SQL quarantine table, extracts DSP,
#   timestamp, and SMI fields when possible, regenerates BodyBinary, and ALWAYS
#   updates each row with Status='Parsed' — even if parsing fails.
#
# HELPER FUNCTIONS:
#   New-SqlConnection - Opens and returns an active SQL connection.
#
# CHANGE LOG:
#   v1.0 - Initial cleanup and formatting.
#   v1.1 - Added logic-explanation flowerboxes.
#   v1.2 - Ensured all Pending rows are updated.
#   v1.3 - Removed continue, added fallback DSP/SMI, fixed missing parameters.
###############################################################################

$version = "v1.3";

# ---------------------------------------------------------
# Initialize counters and open SQL connection
# ---------------------------------------------------------
$parsedCount = 0
$sqlConn = New-SqlConnection

# ---------------------------------------------------------
# Pull all Pending rows (trimmed to avoid hidden chars)
# ---------------------------------------------------------
$cmd = $sqlConn.CreateCommand()
$cmd.CommandText = @"
SELECT Id, BodyText
FROM dbo.MsmqMessageQuarantine
WHERE LTRIM(RTRIM(Status)) = 'Pending'
ORDER BY Id;
"@

$adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
$dt = New-Object System.Data.DataTable
$adapter.Fill($dt) | Out-Null

# ---------------------------------------------------------
# Prepare timestamp components
# ---------------------------------------------------------
$millis = 0
$currentYear  = (Get-Date).Year
$currentMonth = (Get-Date).Month

Write-Host "Parsing DSP / DayTime / SMI fields... $version" -ForegroundColor Green

###############################################################################
# PROCESS EACH ROW
###############################################################################
foreach ($row in $dt.Rows) {

  $body = $row.BodyText
  if (-not $body) { $body = "" }

  # ---------------------------------------------------------
  # Default fallback values (used when parsing fails)
  # ---------------------------------------------------------
  $DSP = "***"
  $SMI = "*D*"
  $timestamp = $null

  # ---------------------------------------------------------
  # Detect QU or XML root
  # ---------------------------------------------------------
  $idxQU  = $body.IndexOf("QU ")
  $idxXML = $body.IndexOf("<?xml ")

  $validIndexes = @()
  if ($idxQU  -ge 0) { $validIndexes += $idxQU }
  if ($idxXML -ge 0) { $validIndexes += $idxXML }

  if ($validIndexes.Count -gt 0) {

    # ---------------------------------------------------------
    # Trim body to start at message root
    # ---------------------------------------------------------
    $startIndex = ($validIndexes | Measure-Object -Minimum).Minimum
    $body = $body.Substring($startIndex)

    # ---------------------------------------------------------
    # Split into lines
    # ---------------------------------------------------------
    $lines = $body -split "`r?`n"

    #############################################################################
    # EXTRACT DSP + TIME FIELD
    #############################################################################
    for ($i = 0; $i -lt $lines.Count; $i++) {

      if ($lines[$i] -match "^\.(\w{7})\s+(\d{6})") {

        $DSP = $Matches[1]
        $timeField = $Matches[2]

        $day    = [int]$timeField.Substring(0,2)
        $hour   = [int]$timeField.Substring(2,2)
        $minute = [int]$timeField.Substring(4,2)

        $timestamp = Get-Date -Year $currentYear -Month $currentMonth -Day $day `
                              -Hour $hour -Minute $minute -Second 0 -Millisecond $millis

        $millis++
        if ($millis -gt 900) { $millis = 0 }

        # ---------------------------------------------------------
        # Extract SMI from next line if available
        # ---------------------------------------------------------
        if ($i + 1 -lt $lines.Count) {
          $SMI = $lines[$i+1].Substring(1,3)
        }

        break
      }
    }
  }
  else {
    Write-Host "ID $($row.Id): No QU/XML root found — forcing Parsed status" -ForegroundColor Yellow
  }

  #############################################################################
  # REGENERATE BODY BINARY
  #############################################################################
  $bodyBinary = [System.Text.Encoding]::UTF8.GetBytes($body)

  #############################################################################
  # UPDATE SQL ROW (ALWAYS)
  #############################################################################
  $update = $sqlConn.CreateCommand()
  $update.CommandText = @"
UPDATE dbo.MsmqMessageQuarantine
SET BodyText        = @BodyText,
    BodyBinary      = @BodyBinary,
    SMI             = @SMI,
    DSP             = @DSP,
    ParsedTimestamp = @TS,
    Status          = 'Parsed'
WHERE Id = @Id;
"@

  # Add parameters
  $update.Parameters.AddWithValue("@BodyText", $body) | Out-Null
  $update.Parameters.Add("@BodyBinary", [System.Data.SqlDbType]::VarBinary, $bodyBinary.Length).Value = $bodyBinary
  $update.Parameters.AddWithValue("@SMI", $SMI) | Out-Null
  $update.Parameters.AddWithValue("@DSP", $DSP) | Out-Null

  if ($timestamp) {
    $update.Parameters.AddWithValue("@TS", $timestamp) | Out-Null
  }
  else {
    $update.Parameters.AddWithValue("@TS", [DBNull]::Value) | Out-Null
  }

  $update.Parameters.AddWithValue("@Id", $row.Id) | Out-Null

  # Execute update
  $update.ExecuteNonQuery() | Out-Null
  $parsedCount++
}

# ---------------------------------------------------------
# Close SQL connection
# ---------------------------------------------------------
$sqlConn.Close()

# ---------------------------------------------------------
# Final operator summary
# ---------------------------------------------------------
Write-Host "Updated $parsedCount rows to Status='Parsed'." -ForegroundColor Cyan

###############################################################################
# END OF SCRIPT
###############################################################################
