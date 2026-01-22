# ****************************************************************************************
# * ASP_getASPDataCenterInfo
# * --------------------------------------------------------------------------------------
# * Author John Buxton
# * --------------------------------------------------------------------------------------
# * Input Parameter: N/A
# * --------------------------------------------------------------------------------------
# * The purpose of this script is to return the data center information 
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
  # * INITIALIZE THE STATUS REPORT
  # * --------------------------------------------------------------------------------------
  [string] $current_server_name = & hostname.exe
   switch ($current_server_name) {
      # * --------------------------------------------------------------------------------------
      # * AIRCOM SERVER TEST SERVERS
      # * --------------------------------------------------------------------------------------
      "pssappt6984"{
       $Airline = "<font color=#B21807>A</font><font color=#0000ff>AL</font>";
       $Environment = "<font size=3 color=#008000><b>Test</font></b>";
       $datacenter = "Skyview 7";
       $from = "AAL ASP Test Skyview 7<aircom@aa.com>";
       $subject = "AAL ASP Test Primary Skyview 7";
    }
      "pssappt6985"{
       $Airline = "<font color=#B21807>A</font><font color=#0000ff>AL</font>";
       $Environment = "<font color=#008000><b>Test</font></b>";
       $datacenter = "Skyview 7";
       $from = "AAL ASP Test Skyview 7<aircom@aa.com>";
       $subject = "AAL ASP Test Secondary Skyview 7";
    }
      "pssappt2756"{
       $Airline = "<font color=#0000ff>ENY</font>";
       $Environment = "<font color=#008000><b>Test</font></b>";
       $datacenter = "Skyview 7";
       $from = "ENY ASP Test Skyview 7<aircom@aa.com>";
       $subject = "AAL ASP Test Primary Skyview 7";
    }
      "pssappt6446"{
       $Airline = "<font color=#0000ff>ENY</font>";
       $Environment = "<font color=#008000><b>Test</font></b>";
       $datacenter = "Skyview 7";
       $from = "ENY ASP Test Skyview 7<aircom@aa.com>";
       $subject = "AAL ASP Test Secondary Skyview 7";
    }
      # * --------------------------------------------------------------------------------------
      # * AIRCOM SERVER STAGE SERVERS
      # * --------------------------------------------------------------------------------------
      "actulmaps400"{
       $Airline = "<font color=#B21807>A</font><font color=#0000ff>AL</font>";
       $Environment = "<font color=#0000ff><b>Stage</font></b>";
       $datacenter = "CDC";
       $from = "AAL ASP STAGE CDC<aircom@aa.com>";
       $subject = "AAL ASP STAGE Primary CDC";
    }
      "acdalmaps400"{
       $Airline = "<font color=#B21807>A</font><font color=#0000ff>AL</font>";
       $Environment = "<font color=#0000ff><b>Stage</font></b>";
       $datacenter = "PDC";
       $from = "AAL ASP STAGE PDC<aircom@aa.com>";
       $subject = "AAL ASP STAGE Secondary PDC";
    }
      "actulmaps401"{
       $Airline = "<font color=#0000ff>ENY</font>";
       $Environment = "<font color=#0000ff><b>Stage</font></b>";
       $datacenter = "CDC";
       $from = "ENY ASP STAGE CDC<aircom@aa.com>";
       $subject = "AAL ASP STAGE Primary CDC";
    }
      "acdalmaps401"{
       $Airline = "<font color=#0000ff>ENY</font>";
       $Environment = "<font color=#0000ff><b>Stage</font></b>";
       $datacenter = "PDC";
       $from = "ENY ASP STAGE PDC<aircom@aa.com>";
       $subject = "AAL ASP STAGE Secondary PDC";
    }

      # * --------------------------------------------------------------------------------------
      # * AIRCOM SERVER PRODUCTION SERVERS
      # * --------------------------------------------------------------------------------------
      "actulmapp400"{
       $Airline = "<font color=#B21807>A</font><font color=#0000ff>AL</font>";
       $Environment = "<font color=#B21807><b>Production</font></b>";
       $datacenter = "CDC";
       $from = "AAL ASP PROD CDC<aircom@aa.com>";
       $subject = "AAL ASP PROD Primary CDC";
    }
      "acdalmapp400"{
       $Airline = "<font color=#B21807>A</font><font color=#0000ff>AL</font>";
       $Environment = "<font color=#B21807><b>Production</font></b>";
       $datacenter = "PDC";
       $from = "AAL ASP PROD PDC<aircom@aa.com>";
       $subject = "AAL ASP PROD Secondary PDC";
    }
      "actulmapp401"{
       $Airline = "<font color=#0000ff>ENY</font>";
       $Environment = "<font color=#B21807><b>Production</font></b>";
       $datacenter = "CDC";
       $from = "ENY ASP PROD CDC<aircom@aa.com>";
       $subject = "AAL ASP PROD Primary CDC";
    }
      "acdalmapp401"{
       $Airline = "<font color=#0000ff>ENY</font>";
       $Environment = "PROD";
       $datacenter = "<font color=#B21807><b>Production</font></b>";
       $from = "ENY ASP PROD PDC<aircom@aa.com>";
       $subject = "AAL ASP PROD Secondary PDC";
    }
  }

  return "N",$current_server_name, $Airline, $Environment, $datacenter, $from, $subject
}

# * --------------------------------------------------------------------------------------
# * EXCEPTION PROCESSING
# * Return Unsuccessful status code
# * --------------------------------------------------------------------------------------
catch {
  $myErrorMsg = "ERROR* * * in ASP Data Center Info: $($_.Exception.Message)";
  [void]$statusMsg.AppendLine($myErrorMsg);
  [void]$statusMsg.AppendLine("");

  return "Y","UNKNOWN", "UNKNOWN", "UNKNOWN", "UNKNOWN", "UNKNOWN Error<aircom@aa.com>", $myErrorMsg
}

