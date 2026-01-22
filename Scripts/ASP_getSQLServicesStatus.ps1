# ****************************************************************************************
# * ASP_getSQLServicesStatus
# * --------------------------------------------------------------------------------------
# * Author John Buxton
# * --------------------------------------------------------------------------------------
# * Input Parameter: N/A
# * --------------------------------------------------------------------------------------
# * The purpose of this script is to get the SQL Server services status and to report and
# * to start any service that is not running.
# * --------------------------------------------------------------------------------------
# * Return values: Raise Alert Y or N
# *
# * ---------------------------------- History Versions ----------------------------------
# * Version   Date       Owner     Description
# * --------  ---------- --------- -------------------------------------------------------
# * v1.0      2025-03-15 jBuxton   Initial release
# * 
# ****************************************************************************************

# **************************************************************
# * Initialize variables
# **************************************************************
  $myVersion = "v1.0";

try {
  # * --------------------------------------------------------------------------------------
  # *   SERVICES
  # *     3. Obtain the SQL Server Services Status
  # * --------------------------------------------------------------------------------------
  [void]$StatusMsg.AppendLine("<font color=#0020C2><b>MS SQL Server Status " + $myVersion + "---------------------------------------</b></font>");
#  $SqlMsg = Get-Service "*AIRCOM_DB*"  | out-string
  $SqlMsg = Get-Service |Select-Object StartType, Status, Name, DisplayName | Where-Object {$_.DisplayName -like "*AIRCOM_DB*"} | out-string
  $SqlMsg = $SqlMsg.Substring(2,$SqlMsg.Length-6);
  [void]$statusMsg.AppendLine($SqlMsg);
  
  [void]$statusMsg.Replace("Running MSSQL$AIRCOM_DB", "<font color=#008000><b>Running </b></font>MSSQL$AIRCOM_DB")
  [void]$statusMsg.Replace("Stopped MSSQL$AIRCOM_DB", "<font color=#F70D1A><b>Stopped </b></font>MSSQL$AIRCOM_DB")
  
  [void]$statusMsg.Replace("Running SQLAgent$AIRCOM_DB", "<font color=#008000><b>Running </b></font>SQLAgent$AIRCOM_DB")
  [void]$statusMsg.Replace("Stopped SQLAgent$AIRCOM_DB", "<font color=#F70D1A><b>Stopped </b></font>SQLAgent$AIRCOM_DB")
  
  [void]$statusMsg.Replace("Running SQLTELEMETRY$AI", "<font color=#008000><b>Running </b></font>SQLTELEMETRY$AI")
  [void]$statusMsg.Replace("Stopped SQLTELEMETRY$AI", "<font color=#FFA500><b>Stopped </b></font>SQLTELEMETRY$AI")

  # * --------------------------------------------------------------------------------------
  # * Build SMS text message
  # * --------------------------------------------------------------------------------------
  $SqlMsg2 = Get-Service |Select-Object Status, DisplayName | Where-Object {$_.DisplayName -like "*AIRCOM_DB*"} | out-string
  if ($SqlMsg2 -like "*Running SQL Server (AIRCOM_DB)*") {
    $myString2 = "sqlDB-UP; "
  } else {
    $myString2 = "sqlDB-DN; "
  }
  if ($SqlMsg2 -like "*Running SQL Server Agent (AIRCOM_DB)*") {
    $myString2 = $myString2 + "agent-UP; "
  } else {
    $myString2 = $myString2 + "agent-DN; "
  }
  [void]$smsMsg.AppendLine($myString2);

  $serviceList ="MSSQL`$AIRCOM_DB","SQLAgent`$AIRCOM_DB";
  $alertUserFlag = "N"
    
  foreach ($service in $serviceList) {
    $myServiceStartType = (Get-Service $service).StartType;
    $myServiceStatus = (Get-Service $service).Status;

    # * --------------------------------------------------------------------------------------
    # *     5. Check each service and START it if it is not running
    # * --------------------------------------------------------------------------------------
    if ($myServiceStatus -ne "Running"  -and $myServiceStartType -ne "Disabled") {
      $myServiceStatus = (Get-Service $service).Status;
      [void]$statusMsg.AppendLine("The $service service is $myServiceStatus. Attempting Restart!");
      [void]$statusMsg.AppendLine("Starting attempt: ...");
      Set-Service -Name $service -StartupType Automatic;
      Start-Service $service -verbose 
      $myServiceStatus = (Get-Service $service).Status;
      [void]$statusMsg.AppendLine("<font color=#F70D1A>$service</font>  <font color=#347C17>$myServiceStatus<</font>");
      [void]$statusMsg.AppendLine("");
      [void]$statusMsg.Replace("color=#0020C2><b>MS SQL Server Status", "color=#F70D1A><b>MS SQL Server Status");
      return "Y";
    }     
  }  

  return "N";
}

# * --------------------------------------------------------------------------------------
# * EXCEPTION PROCESSING
# * Return Unsuccessful status code
# * --------------------------------------------------------------------------------------
catch {
  $myErrorMsg = "ERROR* * * in MS SQL Serivce Status: $($_.Exception.Message)";
  [void]$statusMsg.AppendLine($myErrorMsg);
  [void]$statusMsg.AppendLine("");
 
  return "Y";
}

