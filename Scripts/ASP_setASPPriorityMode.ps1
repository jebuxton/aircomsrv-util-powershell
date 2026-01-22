# ****************************************************************************************
# * ASP_setASPPriorityMode
# * --------------------------------------------------------------------------------------
# * Author John Buxton
# * --------------------------------------------------------------------------------------
# * Input Parameter: N/A
# * --------------------------------------------------------------------------------------
# * The purpose of this script is to detemine if ASP should be put into Priority Mode or
# * taken out of Priority Mode
# * * * * * * * * * * * PRIORITY MODE * * * * * * * * * *
# *
# * Determine if Priority Mode is needed
# * Get QueueCount
# * Get QueueAlertStatus value
# * IF QueueAlertStatus -eq "N" -and QueueCount is -ge 3,000
# *       Set QueueAlertStatus -eq "Y" and Save static
# *       Put MessageProcessor into PRIORITY mode
# *
# * IF QueueAlertStatus -eq "Y" -and QueueCount is -le 500
# *       Set QueueAlertStatus -eq "N" and Save static
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
  # * * * * * * * * * * * PRIORITY MODE * * * * * * * * * *
  # * Determine if Priority Mode is needed
  # * Get QueueCount
  # * Get QueueAlertStatus value
  # * IF QueueAlertStatus -eq "N" -and QueueCount is -ge 3,000
  # *       Set QueueAlertStatus -eq "Y" and Save static
  # *       Disable ENY and DEV TCP users for distribution
  # *       Put MessageProcessor into PRIORITY mode
  # * --------------------------------------------------------------------------------------
  $myQueueCount = (Get-MsmqQueue | ? { $_.QueueName -eq 'private$\fromnetworkq' }).MessageCount
  $QueueAlertStatus = import-clixml -Path D:\PS_Scripts\QueueAlertStatus.xml

  if ($QueueAlertStatus -eq "N" -and $myQueueCount -ge 3000) {
    $QueueAlertStatus = "Y";
    $QueueAlertStatus | export-clixml -path D:\PS_Scripts\QueueAlertStatus.xml
    # * --------------------------------------------------------------------------------------
    # * * * * * * * * * * * PRIORITY MODE * * * * * * * * * *
    # * Call script to set TCP cascade users to inactive for distribution
    # * Restart Message Processor ('MSGPROC') so that Aircom will only process backlog of messages
    # * --------------------------------------------------------------------------------------
    REStart-Service "MSGPROC" -verbose
    [void]$StatusMsg.AppendLine("");
    [void]$StatusMsg.AppendLine("Recovery Mode Started " + $myVersion + "--------------------------------------");
    $statusMsg.Replace("Application Status Monitor", "Application Status Monitor Recovery Mode Started!"); 
    return "Y";
  }

  # * --------------------------------------------------------------------------------------
  # * * * * * * * * * * * PRIORITY MODE * * * * * * * * * *
  # * IF QueueAlertStatus -eq "Y" and QueueCount is -le 500
  # *       Set QueueAlertStatus -eq "N" -and Save static
  # *       Enable ENY and DEV TCP users for distribution
  # * --------------------------------------------------------------------------------------
  if ($QueueAlertStatus -eq "Y" -and $myQueueCount -le 500) {
    $QueueAlertStatus = "N";
    $QueueAlertStatus | export-clixml -path D:\PS_Scripts\QueueAlertStatus.xml
    # * --------------------------------------------------------------------------------------
    # * * * * * * * * * * * PRIORITY MODE * * * * * * * * * *
    # * Call script to set TCP cascade users to ACTIVE for distribution
    # * --------------------------------------------------------------------------------------
    [void]$statusMsg.AppendLine(" ");
    [void]$StatusMsg.AppendLine("Recovery Mode Ended " + $myVersion + "----------------------------------------");
    $statusMsg.Replace("Application Status Monitor", "Application Status Monitor Recovery Mode Completed"); 
    return "N";
  }

  return "N";
}

# * --------------------------------------------------------------------------------------
# * EXCEPTION PROCESSING
# * Return Unsuccessful status code
# * --------------------------------------------------------------------------------------
catch {
  $myErrorMsg = "ERROR* * * in Priority Mode: $($_.Exception.Message)";
  [void]$statusMsg.AppendLine($myErrorMsg);
  [void]$statusMsg.AppendLine("");
  
  return "Y";
}

