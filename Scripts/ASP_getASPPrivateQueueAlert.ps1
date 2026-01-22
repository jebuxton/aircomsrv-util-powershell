# ****************************************************************************************
# * ASP_getASPPrivateQueueAlert
# * --------------------------------------------------------------------------------------
# * Author John Buxton
# * --------------------------------------------------------------------------------------
# * Input Parameter: N/A
# * --------------------------------------------------------------------------------------
# * The purpose of this script is to get the Aircom Server Private Queue Status
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
# * --------------------------------------------------------------------------------------
# * Initialize local 
# * --------------------------------------------------------------------------------------
$myVersion = "v1.0";

try {
  # * --------------------------------------------------------------------------------------
  # * Check fromnetworkq queue value, alert if greater than 1,500 records waiting in the que
  # * Set $alertUserFlag to Yes so that an alert email will be sent
  # * --------------------------------------------------------------------------------------
  $myQueueCount = (Get-MsmqQueue | ? { $_.QueueName -eq 'private$\fromnetworkq' }).MessageCount | out-string

  if ([int]$myQueueCount -ge 1500) {
    [void]$StatusMsg.AppendLine(" ");
    [void]$StatusMsg.AppendLine("<font color=#F70D1A><b>Private Queue Alert * * * " + $myVersion + "---------------------------------</b></font>");
    [void]$StatusMsg.AppendLine(" ");
    [void]$StatusMsg.AppendLine("Queue  Name: fromnetworkq");
    [void]$StatusMsg.AppendLine("Queue Limit: 1500");
    [void]$StatusMsg.AppendLine("Queue Count: <font color=#F70D1A><b>" + $myQueueCount + "</b></font>");
    [void]$statusMsg.Replace("color=#0020C2><b>Private Queue Counts", "color=#F70D1A><b>Private Queue Counts");
    # * --------------------------------------------------------------------------------------
    # * Build SMS text message
    # * --------------------------------------------------------------------------------------
    $myString2 = "*ALERT* workQ-" + $myQueueCount;
    [void]$smsMsg.AppendLine($myString2);

    return "Y";
  }
  # * --------------------------------------------------------------------------------------
  # * Build SMS text message
  # * --------------------------------------------------------------------------------------
  $myString2 = "workQ-" + $myQueueCount;
  [void]$smsMsg.AppendLine($myString2);

  return "N";
}

# * --------------------------------------------------------------------------------------
# * EXCEPTION PROCESSING
# * Return Unsuccessful status code
# * --------------------------------------------------------------------------------------
catch {
  $myErrorMsg = "ERROR* * * in Private Queue Alert: $($_.Exception.Message)";
  [void]$statusMsg.AppendLine($myErrorMsg);
  [void]$statusMsg.AppendLine("");
  
  return "Y";
}

