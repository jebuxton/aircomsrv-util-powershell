###############################################################################
# SCRIPT NAME: QU-Display-MessageStatusSummary
#
# AUTHOR:
#   John Buxton
#
# PURPOSE:
#   Retrieves a grouped summary of message counts by Status from the SQL
#   quarantine table and displays the results in a formatted table for the
#   operator.
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

###############################################################################
# INITIALIZATION
###############################################################################

# ---------------------------------------------------------
# Open SQL connection
# ---------------------------------------------------------
$sqlConn = New-SqlConnection


###############################################################################
# QUERY MESSAGE STATUS SUMMARY
###############################################################################

# ---------------------------------------------------------
# Build SQL query to count messages grouped by Status
# ---------------------------------------------------------
$cmd = $sqlConn.CreateCommand()
$cmd.CommandText = @"
SELECT 
    Status,
    COUNT(*) AS MessageCount
FROM dbo.MsmqMessageQuarantine
GROUP BY Status
ORDER BY Status;
"@

# ---------------------------------------------------------
# Execute query and load results into DataTable
# ---------------------------------------------------------
$adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
$dt = New-Object System.Data.DataTable
$adapter.Fill($dt) | Out-Null


###############################################################################
# DISPLAY SUMMARY
###############################################################################

# ---------------------------------------------------------
# Output formatted summary table for operator visibility
# ---------------------------------------------------------
Write-Host ""
Write-Host "Message Status Summary $version" -ForegroundColor Green
Write-Host "-----------------------"
$dt | Format-Table -AutoSize


###############################################################################
# CLEANUP
###############################################################################

# ---------------------------------------------------------
# Close SQL connection
# ---------------------------------------------------------
$sqlConn.Close()


###############################################################################
# END OF SCRIPT
###############################################################################
