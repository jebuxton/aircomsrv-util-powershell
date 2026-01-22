# ****************************************************************************************
# * ASP_getASPDiskStatus
# * --------------------------------------------------------------------------------------
# * Author John Buxton
# * --------------------------------------------------------------------------------------
# * Input Parameter: N/A
# * --------------------------------------------------------------------------------------
# * The purpose of this script is to get the Aircom Server Disk Space Status
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
$percent = 0;
$alertStatus = "N";
try {
  # * --------------------------------------------------------------------------------------
  # * Build Status Heading
  # * --------------------------------------------------------------------------------------
  [void]$statusMsg.AppendLine(" ");
  [void]$statusMsg.AppendLine("<font color=#0020C2><b>Disk Utilization " + $myVersion + "-------------------------------------------</b></font>");

  $myDiskInfo2 = Get-WmiObject -Class Win32_LogicalDisk |
    Select-Object DeviceID,
    @{Name="% Used";Expression={"{0:P0}" -f (1.00 - ($_.FreeSpace / $_.Size))}} |
    Format-Table -AutoSize   | out-string
    $myDiskInfo2 = $myDiskInfo2.Substring(46,$myDiskInfo2.Length-46)
    $myDiskInfo2 = $myDiskInfo2.Replace("       "," ")
    $myDiskInfo2 = $myDiskInfo2.Replace("E:       
","")
  [void]$smsMsg.AppendLine($myDiskInfo2);


  $myDiskInfo = Get-WmiObject Win32_LogicalDisk | Select-Object DeviceID, VolumeName, @{Name="Size (GB)";Expression={$_.Size/1GB -as [int]}}, @{Name="FreeSpace (GB)";Expression={$_.FreeSpace/1GB -as [int]}}, @{Name="% Used";Expression={"{0:P0}" -f (1.00 - ($_.FreeSpace / $_.Size))}} | Format-Table -AutoSize  | out-string
  $array = $myDiskInfo.Split("`r`n")

  foreach ($element in $array) {
    if ($element -ne "" -and $element.Substring(9,2) -ne "  ") {
      $pos = $element.IndexOf("%");
      if ($pos -gt 45) {
        if ($pos -eq 48) {
          $percentUsed = $element.Substring(45,3);
        }
        if ($pos -eq 47) {
          $percentUsed = $element.Substring(45,2);
        }
        if ($pos -eq 46) {
          $percentUsed = $element.Substring(45,1);
        }
        $element = $element.Substring(0,45) + "<font color=#347C17><b>" + $percentUsed.ToString() + "%</b></font>";
        $percent = [int]$percentUsed;
        if ($percent -gt 65) {
          $element = $element.Substring(0,45) + "<font color=#F70D1A><b>" + $percentUsed.ToString() + "%</b></font>";
          [void]$statusMsg.Replace("color=#0020C2><b>Disk Utilization", "color=#F70D1A><b>Disk Utilization");
          $alertStatus = "Y";
        }
      }
      [void]$statusMsg.AppendLine($element);
   } 
  }
  
  [void]$statusMsg.AppendLine("");

  return $alertStatus;
}
# * --------------------------------------------------------------------------------------
# * EXCEPTION PROCESSING
# * Return Unsuccessful status code
# * --------------------------------------------------------------------------------------
catch {
  $myErrorMsg = "ERROR* * * in ASP Disk Space Status: $($_.Exception.Message)";
  [void]$statusMsg.AppendLine($myErrorMsg);
  [void]$statusMsg.AppendLine("");

  return "Y";
}

