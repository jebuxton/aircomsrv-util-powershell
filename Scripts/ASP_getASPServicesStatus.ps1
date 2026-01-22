# ****************************************************************************************
# * ASP_getASPServicesStatus
# * --------------------------------------------------------------------------------------
# * Author John Buxton
# * --------------------------------------------------------------------------------------
# * Input Parameter: N/A
# * --------------------------------------------------------------------------------------
# * The purpose of this script is to get the ASP services status and to report and
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
  $returnStatus = "N";

try {
  # * --------------------------------------------------------------------------------------
  # * Build Status Heading
  # * --------------------------------------------------------------------------------------
  [void]$statusMsg.AppendLine("<font color=#0020C2><b>Aircom Services Status " + $myVersion + "-------------------------------------</b></font>");
  $myString = Get-Service |Select-Object StartType, Status, Name, DisplayName | Where-Object {$_.DisplayName -like "AS*"} | out-string
  $myString = $myString.Substring(2,$myString.Length-6);
  [void]$statusMsg.AppendLine($myString);
  
  [void]$statusMsg.Replace("Running ASMONITORING", "<font color=#008000><b>Running</b></font> ASMONITORING")
  [void]$statusMsg.Replace("Stopped ASMONITORING", "<font color=#F70D1A><b>Stopped</b></font> ASMONITORING")
  
  [void]$statusMsg.Replace("Running BATAPCNX", "<font color=#008000><b>Running</b></font> BATAPCNX")
  [void]$statusMsg.Replace("Stopped BATAPCNX", "<font color=#F70D1A><b>Stopped</b></font> BATAPCNX")
  
  [void]$statusMsg.Replace("Running DBCNX", "<font color=#008000><b>Running</b></font> DBCNX")
  [void]$statusMsg.Replace("Stopped DBCNX", "<font color=#F70D1A><b>Stopped</b></font> DBCNX")
  
  [void]$statusMsg.Replace("Running EMAILCNX", "<font color=#008000><b>Running</b></font> EMAILCNX")
  [void]$statusMsg.Replace("Stopped EMAILCNX", "<font color=#F70D1A><b>Stopped</b></font> EMAILCNX")
  
  [void]$statusMsg.Replace("Running FILECNX", "<font color=#008000><b>Running</b></font> FILECNX")
  [void]$statusMsg.Replace("Stopped FILECNX", "<font color=#F70D1A><b>Stopped</b></font> FILECNX")
  
  [void]$statusMsg.Replace("Running FRACNX", "<font color=#008000><b>Running</b></font> FRACNX")
  [void]$statusMsg.Replace("Stopped FRACNX", "<font color=#FFA500><b>Stopped</b></font> FRACNX")
  
  [void]$statusMsg.Replace("Running HTTPCNX", "<font color=#008000><b>Running</b></font> HTTPCNX")
  [void]$statusMsg.Replace("Stopped HTTPCNX", "<font color=#FFA500><b>Stopped</b></font> HTTPCNX")
  
  [void]$statusMsg.Replace("Running MQSCNX", "<font color=#008000><b>Running</b></font> MQSCNX")
  [void]$statusMsg.Replace("Stopped MQSCNX", "<font color=#F70D1A><b>Stopped</b></font> MQSCNX")
  
  [void]$statusMsg.Replace("Running MSGPROC", "<font color=#008000><b>Running</b></font> MSGPROC")
  [void]$statusMsg.Replace("Stopped MSGPROC", "<font color=#F70D1A><b>Stopped</b></font> MSGPROC")
  
  [void]$statusMsg.Replace("Running RELAY", "<font color=#008000><b>Running</b></font> RELAY")
  [void]$statusMsg.Replace("Stopped RELAY", "<font color=#F70D1A><b>Stopped</b></font> RELAY")
  
  [void]$statusMsg.Replace("Running SERVICE", "<font color=#008000><b>Running</b></font> SERVICE")
  [void]$statusMsg.Replace("Stopped SERVICE", "<font color=#F70D1A><b>Stopped</b></font> SERVICE")
  
  [void]$statusMsg.Replace("Running TCPCNX", "<font color=#008000><b>Running</b></font> TCPCNX")
  [void]$statusMsg.Replace("Stopped TCPCNX", "<font color=#F70D1A><b>Stopped</b></font> TCPCNX")

  # * --------------------------------------------------------------------------------------
  # * Build SMS text message
  # * --------------------------------------------------------------------------------------
  $myString3 = "";
  $myString2 = Get-Service |Select-Object Status, DisplayName | Where-Object {$_.DisplayName -like "AS*"} | out-string
  $myString2 = $myString2.Substring(70,$myString2.Length-70);
  $myString2 = $myString2.Replace(" Running AS "," ;Running AS ")
  if ($myString2 -like "*Running AS Monitoring Service*") {
    $myString3 = $myString3 + "monitor-UP; "
  } else {
    $myString3 = $myString3 + "monitor-DN; "
  }
  if ($myString2 -like "*Running AS BATAP Connector*") {
    $myString3 = $myString3 + "batap-UP; "
  } else {
    $myString3 = $myString3 + "batap-DN; "
  }
  if ($myString2 -like "*Running AS Database Connector*") {
    $myString3 = $myString3 + "db-UP; "
  } else {
    $myString3 = $myString3 + "db-DN; "
  }
  if ($myString2 -like "*Running AS Email Connector*") {
    $myString3 = $myString3 + "eMail-UP; "
  } else {
    $myString3 = $myString3 + "eMail-DN; "
  }
  if ($myString2 -like "*Running AS File/Printer Connector*") {
    $myString3 = $myString3 + "file-UP; "
  } else {
    $myString3 = $myString3 + "file-DN; "
  }
  if ($myString2 -like "*Running AS IBM MQ Connector *") {
    $myString3 = $myString3 + "mq-UP; "
  } else {
    $myString3 = $myString3 + "mq-DN; "
  }
  if ($myString2 -like "*Running AS Message Processor*") {
    $myString3 = $myString3 + "msg-UP; "
  } else {
    $myString3 = $myString3 + "msg-DN; "
  }
  if ($myString2 -like "*Running AS Communication Relay*") {
    $myString3 = $myString3 + "relay-UP; "
  } else {
    $myString3 = $myString3 + "relay-DN; "
  }
  if ($myString2 -like "*Running AS Base Service*") {
    $myString3 = $myString3 + "base-UP; "
  } else {
    $myString3 = $myString3 + "base-DN; "
  }
  if ($myString2 -like "*Running AS TCP Connector*") {
    $myString3 = $myString3 + "tcp-UP; "
  } else {
    $myString3 = $myString3 + "tcp-DN; "
  }

  [void]$smsMsg.AppendLine($myString3);
 
  $serviceList ="ASMONITORING","BATAPCNX","DBCNX","EMAILCNX","FILECNX","FRACNX","HTTPCNX","MQSCNX","MSGPROC","RELAY","SERVICE","TCPCNX";
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
      [void]$statusMsg.AppendLine("<font color=#F70D1A>$service</font>  <font color=#347C17>$myServiceStatus</font>");
      [void]$statusMsg.AppendLine("");
      [void]$statusMsg.Replace("color=#0020C2><b>Aircom Services Status", "color=#F70D1A><b>Aircom Services Status");
      $returnStatus = "Y";
    }
  }
  
  return $returnStatus;
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

