
# ****************************************************************************************
# * ASPMonitor
# * --------------------------------------------------------------------------------------
# * @author John Buxton
# * @since 2017-09-01
# * --------------------------------------------------------------------------------------
# * @parameter(Input)
# *   Request : Monitor - Automatic request to verify and report any issues.
# *             Report - Generate Status Report and email to subscribers.
# * --------------------------------------------------------------------------------------
# * The purpose of this script is to monitor the Aircom Server services/drive space and  
# * report on their health when there is an issue and when possible try to correct the
# * issue by restarting the failing service(s).
# * --------------------------------------------------------------------------------------
# * BUILD THE STATUS REPORT
# * --------------------------------------------------------------------------------------
# *   DRIVE SPACE
# *     1. Obtain a list of all the drives for the status report
# * --------------------------------------------------------------------------------------
# *   SERVICES
# *     2. Obtain the Aircom Server Services status
# *     3. Obtain the SQL Server Services status
# * --------------------------------------------------------------------------------------
# *   EMAIL STATUS REPORT   
# *     4. Build the email message, and email to subscribers. 
# * --------------------------------------------------------------------------------------
# * @return
# *   Status Report - Request=Monitor, email Status Report if any issues are detected
# *                   Request=Report,  email Status Report
# *
# * ---------------------------------- History Versions ----------------------------------
# * Version   Date       Owner     Description
# * --------  ---------- --------- -------------------------------------------------------
# * v1.0      2025-03-15 jBuxton   Initial release
# * 
# ****************************************************************************************
param (
    [string]$Request) "Report"

try {
  # * --------------------------------------------------------------------------------------
  # * Initialize LOCAL VARIABLES
  # * --------------------------------------------------------------------------------------
  $myVersion = "v1.0";
  $alertUserFlag = "";
  $statusMsg = New-Object System.Text.StringBuilder;
  $smsMsg = New-Object System.Text.StringBuilder;
  $tempStatusMsg = New-Object System.Text.StringBuilder;
  $CurrentDateTime = Get-Date -Format "dd-MMM-yy HH:mm"
  $zTime = (Get-Date).ToUniversalTime().ToString("dd-MMM-yy HH:mm")

  # * --------------------------------------------------------------------------------------
  # * Get Data Center Information and Email address and Subject
  # * --------------------------------------------------------------------------------------
  $result = (& "D:\PS_Scripts\ASP_getASPDataCenterInfo.ps1");
  $alertUserFlag = $alertUserFlag + $($result[0]);
  $current_server_name            = $($result[1]);
  $Airline                        = $($result[2]);
  $Environment                    = $($result[3]);
  $datacenter                     = $($result[4]);
  $from                           = $($result[5]);
  $subject                        = $($result[6]);

  [void]$statusMsg.AppendLine("<font face='Courier' color=#000000><PRE>");
  [void]$statusMsg.AppendLine("<font color=#347C17><b><h2>Aircom Server Application Health " + $Request + "</h2></b></font>");
  [void]$statusMsg.AppendLine("  Date Time: <b>" + $CurrentDateTime + " (" + $zTime + "z)</b>");
  [void]$statusMsg.AppendLine("    Airline: <b>" + $airline + "</b>");
  [void]$statusMsg.AppendLine("Environment: <b>" + $environment + "</b>");
  [void]$statusMsg.AppendLine(" Datacenter: <b>" + $datacenter + "</b>");
  [void]$statusMsg.AppendLine("Server Name: <b><font color=#0020C2>" + $current_server_name + "</font></b>");
  [void]$statusMsg.AppendLine("    Version: <b>" + $myVersion + "</b>");

  [void]$smsMsg.AppendLine("Aircom Status " + $Request);
  [void]$smsMsg.AppendLine($zTime + "z");
  [void]$smsMsg.AppendLine($airline);
  [void]$smsMsg.AppendLine($environment + ": " + $datacenter);
  [void]$smsMsg.AppendLine($current_server_name);
  [void]$smsMsg.AppendLine($myVersion);


  # * --------------------------------------------------------------------------------------
  # * BUILD THE STATUS REPORT
  # * --------------------------------------------------------------------------------------
  # *   DRIVE SPACE
  # *     1. Obtain a list of all the drives for the status report
  # *   SERVICES
  # *     2. Obtain the Aircom Server Services Status 
  # *     3. Obtain the SQL Server Services Status
  # *   QUEUES
  # *     4a. Obtain the ASP Private Queue Status
  # *     4b. Alert if 'private$\fromnetworkq' is greater than 1,000
  # *     5.  Determine if ASP should be placed into Priority Mode or removed from Priority 
  # * --------------------------------------------------------------------------------------
  $alertUserFlag = "";
  $alertUserFlag = $alertUserFlag + (& "D:\PS_Scripts\ASP_getASPDiskStatus.ps1");
  $alertUserFlag = $alertUserFlag + (& "D:\PS_Scripts\ASP_getASPServicesStatus.ps1");
  $alertUserFlag = $alertUserFlag + (& "D:\PS_Scripts\ASP_getSQLServicesStatus.ps1");
  $alertUserFlag = $alertUserFlag + (& "D:\PS_Scripts\ASP_checkApplicationStatus.ps1");
  $alertUserFlag = $alertUserFlag + (& "D:\PS_Scripts\ASP_getAircomLogCount.ps1");

  $alertUserFlag = $alertUserFlag + (& "D:\PS_Scripts\ASP_getASPPrivateQueueStatus.ps1");
  $alertUserFlag = $alertUserFlag + (& "D:\PS_Scripts\ASP_getASPPrivateQueueAlert.ps1");
  $alertUserFlag = $alertUserFlag + (& "D:\PS_Scripts\ASP_setASPPriorityMode.ps1");

  # * --------------------------------------------------------------------------------------
  # *   Build Distrbution lists   
  # *     6a. build eMail distribution list
  # *         build phone alert distribution list 
  # * --------------------------------------------------------------------------------------
  $eMailList = "Johnny B<John.Buxton@aa.com>, Skunk Works<John.Raab@aa.com>, Salik<Salik.Shariff@aa.com>, Ram<Ram.Lokireddy@aa.com>";
  #$eMailList = "Johnny B<John.Buxton@aa.com>";
  #$smsList = "<2052769724@tmomail.net>, <8176915825@tmomail.net>";
  #$smsList = "<2052769724@tmomail.net>";
  $smsList = "";

  [void]$StatusMsg.AppendLine(" ");
  [void]$StatusMsg.AppendLine("<font color=#0020C2><b>eMail Distribution " + $myVersion + "-----------------------------------------</b></font>");
  [void]$StatusMsg.AppendLine("eMail List: " + $eMailList);
  [void]$StatusMsg.AppendLine("      From: " + $from);
  [void]$StatusMsg.AppendLine("   Subject: " + $subject);

  if ($alertUserFlag.IndexOf("Y") -gt 0) {
    $alertUserFlag = "Y"
  }
  else {
    $alertUserFlag = "N"
  }

  # * --------------------------------------------------------------------------------------
  # *   EMAIL STATUS REPORT   
  # *     6b. If flag to send an email notice is set or a status report request is being made,
  # *        then build the email message, and email to subscribers. 
  # * --------------------------------------------------------------------------------------
  if ($alertUserFlag -eq "Y") {
    $lastReportTime = import-clixml -Path D:\PS_Scripts\ASP_lastReportTime.xml
    $currentReportTime = Get-Date 
    $timeDifference = New-TimeSpan -Start $lastReportTime -End $currentReportTime
    $minutesDifference = $($timeDifference.TotalMinutes)
    if ($minutesDifference -ge 15) {
      [void]$statusMsg.Replace("color=#347C17><b><h2>Aircom Server Application Status", "color=#F70D1A><b><h2>Aircom Server Application Status");
      $currentReportTime | export-clixml -path D:\PS_Scripts\LastReportTime.xml
      [void]$statusMsg.AppendLine("eMail Sent: ASP Status " + $Request + " " + $myVersion + "</PRE></font>");
      [void]$statusMsg.Replace("`r`n", "<br/>");
      # * --------------------------------------------------------------------------------------
      # Build HTML Email
      # * --------------------------------------------------------------------------------------
      $to      =  $eMailList 
      $from    =  $from
      $subject =  "*** ALERT! *** ASP Status Report "
      $body    =  $statusMsg             
      # * --------------------------------------------------------------------------------------
      # Create .Net Email Object
      # Create an SMTP Server Object
      # * --------------------------------------------------------------------------------------
      $mail = New-Object System.Net.Mail.Mailmessage $from, $to, $subject, $body
      $mail.IsBodyHTML=$true
      $server = "smtp.qcorpaa.aa.com"
      $port   = 25
      $Smtp = New-Object System.Net.Mail.SMTPClient $server,$port
      $Smtp.Credentials = [system.Net.CredentialCache]::DefaultNetworkCredentials
      # * --------------------------------------------------------------------------------------
      # Send .Net Email Object
      # * --------------------------------------------------------------------------------------
      $smtp.send($mail)
    }
  }

  if ($alertUserFlag -eq "N" -And $Request -eq "Report") {
    [void]$statusMsg.AppendLine("eMail Sent: ASP Status " + $Request + " " + $myVersion + "</PRE></font>");
    [void]$statusMsg.Replace("`r`n", "<br/>");
    # * --------------------------------------------------------------------------------------
    # Build HTML Email
    # * --------------------------------------------------------------------------------------
      $to      =  $eMailList 
      $from    =  $from
      $subject =  "ASP Status Report"
      $body    =  $statusMsg             
    # * --------------------------------------------------------------------------------------
    # Create .Net Email Object
    # Create an SMTP Server Object
    # * --------------------------------------------------------------------------------------
      $mail = New-Object System.Net.Mail.Mailmessage $from, $to, $subject, $body
      $mail.IsBodyHTML=$true
      $server = "smtp.qcorpaa.aa.com"
      $port   = 25
      $Smtp = New-Object System.Net.Mail.SMTPClient $server,$port
      $Smtp.Credentials = [system.Net.CredentialCache]::DefaultNetworkCredentials
    # * --------------------------------------------------------------------------------------
    # Send .Net Email Object
    # * --------------------------------------------------------------------------------------
      $smtp.send($mail)
  }

  Write-Host $statusMsg

  if ($alertUserFlag -eq "Y" -and $smsList -ne "") {
    # * --------------------------------------------------------------------------------------
    # Build SMS Email
    # * --------------------------------------------------------------------------------------
      $to      =  $smsList 
      $from    =  $from
      $subject =  "ASP Alert!"
      $body    =  $smsMsg             
    # * --------------------------------------------------------------------------------------
    # Create .Net Email Object
    # Create an SMTP Server Object
    # * --------------------------------------------------------------------------------------
      $mail2 = New-Object System.Net.Mail.Mailmessage $from, $to, $subject, $body
      $mail2.IsBodyHTML=$true
      $server = "smtp.qcorpaa.aa.com"
      $port   = 25
      $Smtp2 = New-Object System.Net.Mail.SMTPClient $server,$port
      $Smtp2.Credentials = [system.Net.CredentialCache]::DefaultNetworkCredentials
    # * --------------------------------------------------------------------------------------
    # Send .Net Email Object
    # * --------------------------------------------------------------------------------------
      $smtp2.send($mail2)
  }


}

# * --------------------------------------------------------------------------------------
# * EXCEPTION PROCESSING
# * Return Unsuccessful status code
# * --------------------------------------------------------------------------------------
catch {
  $myErrorMsg = "ERROR* * * in ASP Monitor " + $Request + " " + $myVersion + " $($_.Exception.Message)";
  [void]$statusMsg.AppendLine("");
  [void]$statusMsg.AppendLine($myErrorMsg);
  [void]$statusMsg.AppendLine("" + "</PRE></font>");
  Write-Host $myErrorMsg -ForegroundColor Red

  $sendMailMessageSplat = @{
    From = $from
    To = "<John.Buxton@aa.com>", "<Glenn.Davis@aa.com>", "<John.Raab@aa.com>", "Salik.Shariff@aa.com", "<2052769724@tmomail.net>"
    Subject = "ASP Status Report ** Exception ** "
    SmtpServer = 'smtp.qcorpaa.aa.com'
    Body=$myErrorMsg
  }
  Send-MailMessage @sendMailMessageSplat
  Write-Host "Sent eMail Line 220";
  Write-Host $statusMsg
}

  
