# ****************************************************************************************
# * ASP_getASPPrivateQueueStatus
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
$myVersion      = "v1.0";
$tempStatusMsg  = New-Object System.Text.StringBuilder;

try {
  # * --------------------------------------------------------------------------------------
  # * Get current private queue counts and include it in the report
  # * --------------------------------------------------------------------------------------
  $myQueueCount = Get-MsmqQueue -QueueType Private | Select-Object QueueName, MessageCount | out-string
  
  [void]$tempStatusMsg.AppendLine("<font color=#0020C2><b>Private Queue Counts " + $myVersion + "---------------------------------------</b></font>");
  [void]$tempStatusMsg.AppendLine($myQueueCount);

  # * --------------------------------------------------------------------------------------
  # * Remove excessive cr lf  (`r`n")
  # * --------------------------------------------------------------------------------------
  $tempStatusMsg = $tempStatusMsg -replace "`r`n`r`n`r`n", "`r`n"
  $tempStatusMsg = $tempStatusMsg -replace "`r`n`r`n", "`r`n"
  $tempStatusMsg = $tempStatusMsg.Substring(0,$tempStatusMsg.Length-6);

  [void]$statusMsg.AppendLine($tempStatusMsg);

  return "N";
}

# * --------------------------------------------------------------------------------------
# * EXCEPTION PROCESSING
# * Return Unsuccessful status code
# * --------------------------------------------------------------------------------------
catch {
  $myErrorMsg = "ERROR* * * in Private Queue Status: $($_.Exception.Message)";
  [void]$statusMsg.AppendLine($myErrorMsg);
  [void]$statusMsg.AppendLine("");
 
  return "Y";
}

