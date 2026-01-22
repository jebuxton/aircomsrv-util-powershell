
# ****************************************************************************************
# * ASP_getAircomLogStatuss
# * --------------------------------------------------------------------------------------
# * Author John Buxton
# * --------------------------------------------------------------------------------------
# * Input Parameter:
# *       n/a
# * --------------------------------------------------------------------------------------
# * The purpose of this script is to check the Aircom Server traffic log to determine 
# *     if Aircom Server is receiving messages. 
# * from Cyberjet and ASP Database messages
# * Red    - > 400   messages in the last 5 minutes - ALERT
# * Orange - < 1,000 messages in the last 5 minutes
# * Green  - > 1,000 messages in the last 5 minutes
# * --------------------------------------------------------------------------------------
# * Return values
# *      No Alert: N
# *         Alert: Y
# *
# * ---------------------------------- History Versions ----------------------------------
# * Version   Date       Owner     Description
# * --------  ---------- --------- -------------------------------------------------------
# * v1.0      2025-11-19 jBuxton   Initial release
# * 
# ****************************************************************************************

try {
  # * --------------------------------------------------------------------------------------
  # * Initialize local and database SQL variables
  # * Need to obtain userID and password fro key valut
  # * --------------------------------------------------------------------------------------
  [string] $SqlServer   = [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName
  [string] $Database    = "AircomSrv"
  [string] $SqlUsername = "aircomusr"
  [string] $SqlPass     = "S!ta2000"
  [string] $myAlertStatus = "Y";

  # **********************************************************
  # * Build connection string with user ID and password
  # **********************************************************
  $connectionString = "Server=$SqlServer;Database=$Database;User ID=$SqlUsername;Password=$SqlPass;"

  # **********************************************************
  # * Open connection
  # **********************************************************
  $SQLconnection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
  $SQLconnection.Open()

  # **********************************************************
  # * Run query
  # **********************************************************
  $sql = @"
  SELECT COUNT(*) AS MsgCount
  FROM Logt
  WHERE LogTTS >= DATEADD(MINUTE, -5, GETUTCDATE());
"@

  $SQLcommand = $SQLconnection.CreateCommand()
  $SQLcommand.CommandText = $sql

  $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $SQLcommand
  $table   = New-Object System.Data.DataTable
  [void]$adapter.Fill($table)

  $SQLconnection.Close()

  $row   = $table.Rows[0]
  $count = [int]$row.MsgCount

  # **********************************************************
  # * Determine color based on thresholds
  # **********************************************************
  if ($count -le 400) {
      $color = "red"
      $myAlertStatus = "Y";  }
  elseif ($count -le 1000) {
          $color = "orange"
          $myAlertStatus = "N";  }
  else {
          $color = "green"
          $myAlertStatus = "N";
  }

  # **********************************************************
  # * Build output message
  # **********************************************************
  $msg = " Aircom: <font color=$color><b>Number of messages within the 5 minutes</b></font>: $count</br>"
  [void]$statusMsg.AppendLine($msg);
  return $myAlertStatus;
}
catch {
  # **********************************************************
  # * Exception Handling
  # **********************************************************
  $errorMsg = " Aircom: ERROR* * * Unable to query Logt table: $($_.Exception.Message)</br>"
  $statusMsg.Remove($statusMsg.Length - 2, 2);
  [void]$statusMsg.AppendLine($errorMsg);
  return "Y";
}

