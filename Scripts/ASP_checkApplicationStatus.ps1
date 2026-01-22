# ****************************************************************************************
# * ASP_checkApplicationStatus
# * --------------------------------------------------------------------------------------
# * Author John Buxton
# * --------------------------------------------------------------------------------------
# * Input Parameter:
# *       n/a
# * --------------------------------------------------------------------------------------
# * The purpose of this script is to check the Cyberjet status by quering these messages 
# * from Cyberjet and ASP Database messages
# * WXR - Weather: Alert if greater than 15 minutes
# * FPR - Flight Progress Report - Alert if greater than 15 minutes
# * FPN - Flight Plan - Alert if greater than 15 minutes
# * Alert if no message from Cyberjet within the last 5 minutes 
# * --------------------------------------------------------------------------------------
# * Active or not active For Distribution TCP Cascade Users
# * --------------------------------------------------------------------------------------
# * Return values
# *      No Alert: N
# *         Alert: Y
# *
# * ---------------------------------- History Versions ----------------------------------
# * Version   Date       Owner     Description
# * --------  ---------- --------- -------------------------------------------------------
# * v1.0      2025-04-09 jBuxton   Initial release
# * 
# ****************************************************************************************


try {
  # * --------------------------------------------------------------------------------------
  # * Initialize local and database SQL variables
  # * Need to obtain userID and password fro key valut
  # * --------------------------------------------------------------------------------------
  [string] $myVersion = "v1.0";
  [string] $alertStatus = "Y";
  [string] $myString = "";
  
  [string] $SqlServer = "pssappt6984.qcorpaa.aa.com";
  [string] $Database = "AircomSrv";
  [string] $SqlUsername = "aircomusr"; 
  [string] $SqlPass = "S!ta2000";

  # **************************************************************
  # * Initialize variables and obtain the SQL credentials
  # **************************************************************
  [void]$statusMsg.AppendLine("<font color=#0020C2><b>Application Message Status " + $myVersion + "---------------------------------</b></font>");
  [void]$statusMsg.AppendLine("App Message Status     Message Date Time   Message and Flt Info");
  [void]$statusMsg.AppendLine("---------------------- ------------------- ---------------------");

  # * --------------------------------------------------------------------------------------
  # * Build connection string for Server, Database, UserId and password
  # * Build SQL query string to update TCP cascade users active for distribution
  # * Create SQL connection object
  # * Open SQL connection object
  # * Create SQL command object using SQL query string and SQL connection object
  # * Execute SQL command object
  # * Save the SQL completion return code value
  # * --------------------------------------------------------------------------------------
  $SQLconnectionString = "Server=$SqlServer;Database=$Database;User ID=$SqlUsername;Password=$SqlPass;"

  $SQLconnection = New-Object System.Data.SqlClient.SqlConnection($SQLconnectionString)
  $SQLconnection.Open()

  # * --------------------------------------------------------------------------------------
  # * Cyberjet: AGM
  # * --------------------------------------------------------------------------------------
  [string] $SqlQuery = "SELECT TOP (1) [LogtTS],[MessageType],[AircraftId]," + 
                                      "[FlightId],[DepartureAirportIATA],"+
                                      "[ArrivalAirportIATA] " +
                                      "FROM [AircomSrv].[dbo].[Logt]  " +
                                      "WHERE [LogtTS] > DATEADD(Minute, -5, GETDATE())  " +
                                      " AND [Direction] = 2 " + 
                                      " AND [MessageType] LIKE '%AGM Messages From Cyberjet%' " + 
                                      "ORDER BY [LogtTS] DESC";

  $SQLcommand = New-Object System.Data.SqlClient.SqlCommand($SQLquery, $SQLconnection)
  $Reader = $SQLcommand.ExecuteReader()
  if ($Reader.HasRows) {
    while ($Reader.Read()) {
      $LogtTS               = $Reader.GetValue(0)
      $MessageType          = $Reader.GetValue(1)
      $AircraftId           = $Reader.GetValue(2)
      $FlightId             = $Reader.GetValue(3)
      $DepartureAirportIATA = $Reader.GetValue(4)
      $ArrivalAirportIATA   = $Reader.GetValue(5)
      $alertStatus = "N";
      $myString = "Cyberjet: <font color=#008000><b>AGM Messages</b></font>" + " " + $LogtTS + " " + $MessageType + " " + $AircraftId + " " + $FlightId + " " + $DepartureAirportIATA + " " + $ArrivalAirportIATA 
      [void]$statusMsg.AppendLine($myString);
    } 
  } else {
      $myString = "Cyberjet: <font color=#F70D1A><b>No AGM Msgs </b></font>"
      [void]$statusMsg.AppendLine($myString);
    }
 
  $SQLcommand.Dispose();
  $Reader.Dispose();
  $Reader.Close();

  # * --------------------------------------------------------------------------------------
  # * Cyberjet: CMD
  # * --------------------------------------------------------------------------------------
  [string] $SqlQuery = "SELECT TOP (1) [LogtTS],[MessageType],[AircraftId]," + 
                                      "[FlightId],[DepartureAirportIATA],"+
                                      "[ArrivalAirportIATA] " +
                                      "FROM [AircomSrv].[dbo].[Logt]  " +
                                      "WHERE [LogtTS] > DATEADD(Minute, -5, GETDATE())  " +
                                      " AND [Direction] = 2 " + 
                                      " AND [MessageType] LIKE '%CMD Message From Cyberjet%' " + 
                                      "ORDER BY [LogtTS] DESC";

  $SQLcommand = New-Object System.Data.SqlClient.SqlCommand($SQLquery, $SQLconnection)
  $Reader = $SQLcommand.ExecuteReader()
  if ($Reader.HasRows) {
    while ($Reader.Read()) {
      $LogtTS               = $Reader.GetValue(0)
      $MessageType          = $Reader.GetValue(1)
      $AircraftId           = $Reader.GetValue(2)
      $FlightId             = $Reader.GetValue(3)
      $DepartureAirportIATA = $Reader.GetValue(4)
      $ArrivalAirportIATA   = $Reader.GetValue(5)
      $alertStatus = "N";
      $myString = "Cyberjet: <font color=#008000><b>CMD Messages</b></font>" + " " + $LogtTS + " " + $MessageType + " " + $AircraftId + " " + $FlightId + " " + $DepartureAirportIATA + " " + $ArrivalAirportIATA 
      [void]$statusMsg.AppendLine($myString);
    } 
  } else {
      $myString = "Cyberjet: <font color=#F70D1A><b>No CMD Msgs </b></font>"
      [void]$statusMsg.AppendLine($myString);
    }
 
  $SQLcommand.Dispose();
  $Reader.Dispose();
  $Reader.Close();

  # * --------------------------------------------------------------------------------------
  # * Cyberjet: Flight Plans
  # * --------------------------------------------------------------------------------------
  [string] $SqlQuery = "SELECT TOP (1) [LogtTS],[MessageType],[AircraftId]," + 
                                      "[FlightId],[DepartureAirportIATA],"+
                                      "[ArrivalAirportIATA] " +
                                      "FROM [AircomSrv].[dbo].[Logt]  " +
                                      "WHERE [LogtTS] > DATEADD(Minute, -5, GETDATE())  " +
                                      " AND [Direction] = 2 " + 
                                      " AND [MessageType] LIKE '%FMC-FPN From Cyberjet%' " + 
                                      "ORDER BY [LogtTS] DESC";

  $SQLcommand = New-Object System.Data.SqlClient.SqlCommand($SQLquery, $SQLconnection)
  $Reader = $SQLcommand.ExecuteReader()
  if ($Reader.HasRows) {
    while ($Reader.Read()) {
      $LogtTS               = $Reader.GetValue(0)
      $MessageType          = $Reader.GetValue(1)
      $AircraftId           = $Reader.GetValue(2)
      $FlightId             = $Reader.GetValue(3)
      $DepartureAirportIATA = $Reader.GetValue(4)
      $ArrivalAirportIATA   = $Reader.GetValue(5)
      $alertStatus = "N";
      $myString = "Cyberjet: <font color=#008000><b>FMC-FPN Msgs</b></font>" + " " + $LogtTS + " " + $MessageType + " " + $AircraftId + " " + $FlightId + " " + $DepartureAirportIATA + " " + $ArrivalAirportIATA 
      [void]$statusMsg.AppendLine($myString);
    } 
  } else {
      $myString = "Cyberjet: <font color=#F70D1A><b>No FMC-FPN </b></font>"
      [void]$statusMsg.AppendLine($myString);
    }
 
  $SQLcommand.Dispose();
  $Reader.Dispose();
  $Reader.Close();

  # * --------------------------------------------------------------------------------------
  # * Cyberjet: Flight Plans Unavailable
  # * --------------------------------------------------------------------------------------
  [string] $SqlQuery = "SELECT TOP (1) [LogtTS],[MessageType],[AircraftId]," + 
                                      "[FlightId],[DepartureAirportIATA],"+
                                      "[ArrivalAirportIATA] " +
                                      "FROM [AircomSrv].[dbo].[Logt]  " +
                                      "WHERE [LogtTS] > DATEADD(Minute, -5, GETDATE())  " +
                                      " AND [Direction] = 2 " + 
                                      " AND [MessageType] LIKE '%Messages From Cyberjet-FP not Available%' " + 
                                      "ORDER BY [LogtTS] DESC";


  $SQLcommand = New-Object System.Data.SqlClient.SqlCommand($SQLquery, $SQLconnection)
  $Reader = $SQLcommand.ExecuteReader()
  if ($Reader.HasRows) {
    while ($Reader.Read()) {
      $LogtTS               = $Reader.GetValue(0)
      $MessageType          = $Reader.GetValue(1)
      $AircraftId           = $Reader.GetValue(2)
      $FlightId             = $Reader.GetValue(3)
      $DepartureAirportIATA = $Reader.GetValue(4)
      $ArrivalAirportIATA   = $Reader.GetValue(5)
      $myString = "Cyberjet: <font color=#F70D1A><b>FPN Unavail </b></font>" + " " + $LogtTS + " " + $MessageType + " " + $AircraftId + " " + $FlightId + " " + $DepartureAirportIATA + " " + $ArrivalAirportIATA 
      [void]$statusMsg.AppendLine($myString);
    } 
  } else {
      $myString = "Cyberjet: <font color=#008000><b>No  FPN Unal </b></font>No 'From Cyberjet-FP not Available' msg has been recevied"
      [void]$statusMsg.AppendLine($myString);
    }
 
  $SQLcommand.Dispose();
  $Reader.Dispose();
  $Reader.Close();

  # * --------------------------------------------------------------------------------------
  # * Cyberjet: LDI
  # * --------------------------------------------------------------------------------------
  [string] $SqlQuery = "SELECT TOP (1) [LogtTS],[MessageType],[AircraftId]," + 
                                      "[FlightId],[DepartureAirportIATA],"+
                                      "[ArrivalAirportIATA] " +
                                      "FROM [AircomSrv].[dbo].[Logt]  " +
                                      "WHERE [LogtTS] > DATEADD(Minute, -5, GETDATE())  " +
                                      " AND [Direction] = 2 " + 
                                      " AND [MessageType] LIKE '%FMC-LDI From Cyberjet%' " + 
                                      "ORDER BY [LogtTS] DESC";

  $SQLcommand = New-Object System.Data.SqlClient.SqlCommand($SQLquery, $SQLconnection)
  $Reader = $SQLcommand.ExecuteReader()
  if ($Reader.HasRows) {
    while ($Reader.Read()) {
      $LogtTS               = $Reader.GetValue(0)
      $MessageType          = $Reader.GetValue(1)
      $AircraftId           = $Reader.GetValue(2)
      $FlightId             = $Reader.GetValue(3)
      $DepartureAirportIATA = $Reader.GetValue(4)
      $ArrivalAirportIATA   = $Reader.GetValue(5)
      $alertStatus = "N";
      $myString = "Cyberjet: <font color=#008000><b>FMC-LDI Msgs</b></font>" + " " + $LogtTS + " " + $MessageType + " " + $AircraftId + " " + $FlightId + " " + $DepartureAirportIATA + " " + $ArrivalAirportIATA 
      [void]$statusMsg.AppendLine($myString);
    } 
  } else {
      $myString = "Cyberjet: <font color=#F70D1A><b>No LDI Msgs </b></font>"
      [void]$statusMsg.AppendLine($myString);
    }
 
  $SQLcommand.Dispose();
  $Reader.Dispose();
  $Reader.Close();

  # * --------------------------------------------------------------------------------------
  # * Cyberjet: PER
  # * --------------------------------------------------------------------------------------
  [string] $SqlQuery = "SELECT TOP (1) [LogtTS],[MessageType],[AircraftId]," + 
                                      "[FlightId],[DepartureAirportIATA],"+
                                      "[ArrivalAirportIATA] " +
                                      "FROM [AircomSrv].[dbo].[Logt]  " +
                                      "WHERE [LogtTS] > DATEADD(Minute, -5, GETDATE())  " +
                                      " AND [Direction] = 2 " + 
                                      " AND [MessageType] LIKE '%FMC-PER From Cyberjet%' " + 
                                      "ORDER BY [LogtTS] DESC";

  $SQLcommand = New-Object System.Data.SqlClient.SqlCommand($SQLquery, $SQLconnection)
  $Reader = $SQLcommand.ExecuteReader()
  if ($Reader.HasRows) {
    while ($Reader.Read()) {
      $LogtTS               = $Reader.GetValue(0)
      $MessageType          = $Reader.GetValue(1)
      $AircraftId           = $Reader.GetValue(2)
      $FlightId             = $Reader.GetValue(3)
      $DepartureAirportIATA = $Reader.GetValue(4)
      $ArrivalAirportIATA   = $Reader.GetValue(5)
      $alertStatus = "N";
      $myString = "Cyberjet: <font color=#008000><b>FMC-PER Msgs</b></font>" + " " + $LogtTS + " " + $MessageType + " " + $AircraftId + " " + $FlightId + " " + $DepartureAirportIATA + " " + $ArrivalAirportIATA 
      [void]$statusMsg.AppendLine($myString);
    } 
  } else {
      $myString = "Cyberjet: <font color=#F70D1A><b>No PER Msgs </b></font>"
      [void]$statusMsg.AppendLine($myString);
    }
 
  $SQLcommand.Dispose();
  $Reader.Dispose();
  $Reader.Close();

  # * --------------------------------------------------------------------------------------
  # * Cyberjet: POS
  # * --------------------------------------------------------------------------------------
  [string] $SqlQuery = "SELECT TOP (1) [LogtTS],[MessageType],[AircraftId]," + 
                                      "[FlightId],[DepartureAirportIATA],"+
                                      "[ArrivalAirportIATA] " +
                                      "FROM [AircomSrv].[dbo].[Logt]  " +
                                      "WHERE [LogtTS] > DATEADD(Minute, -5, GETDATE())  " +
                                      " AND [Direction] = 2 " + 
                                      " AND [MessageType] LIKE '%FMC-POS From Cyberjet%' " + 
                                      "ORDER BY [LogtTS] DESC";

  $SQLcommand = New-Object System.Data.SqlClient.SqlCommand($SQLquery, $SQLconnection)
  $Reader = $SQLcommand.ExecuteReader()
  if ($Reader.HasRows) {
    while ($Reader.Read()) {
      $LogtTS               = $Reader.GetValue(0)
      $MessageType          = $Reader.GetValue(1)
      $AircraftId           = $Reader.GetValue(2)
      $FlightId             = $Reader.GetValue(3)
      $DepartureAirportIATA = $Reader.GetValue(4)
      $ArrivalAirportIATA   = $Reader.GetValue(5)
      $alertStatus = "N";
      $myString = "Cyberjet: <font color=#008000><b>FMC-POS Msgs</b></font>" + " " + $LogtTS + " " + $MessageType + " " + $AircraftId + " " + $FlightId + " " + $DepartureAirportIATA + " " + $ArrivalAirportIATA 
      [void]$statusMsg.AppendLine($myString);
    } 
  } else {
      $myString = "Cyberjet: <font color=#F70D1A><b>No POS Msgs </b></font>"
      [void]$statusMsg.AppendLine($myString);
    }
 
  $SQLcommand.Dispose();
  $Reader.Dispose();
  $Reader.Close();

  # * --------------------------------------------------------------------------------------
  # * Cyberjet: PWI
  # * --------------------------------------------------------------------------------------
  [string] $SqlQuery = "SELECT TOP (1) [LogtTS],[MessageType],[AircraftId]," + 
                                      "[FlightId],[DepartureAirportIATA],"+
                                      "[ArrivalAirportIATA] " +
                                      "FROM [AircomSrv].[dbo].[Logt]  " +
                                      "WHERE [LogtTS] > DATEADD(Minute, -5, GETDATE())  " +
                                      " AND [Direction] = 2 " + 
                                      " AND [MessageType] LIKE '%FMC-PWI From Cyberjet%' " + 
                                      "ORDER BY [LogtTS] DESC";

  $SQLcommand = New-Object System.Data.SqlClient.SqlCommand($SQLquery, $SQLconnection)
  $Reader = $SQLcommand.ExecuteReader()
  if ($Reader.HasRows) {
    while ($Reader.Read()) {
      $LogtTS               = $Reader.GetValue(0)
      $MessageType          = $Reader.GetValue(1)
      $AircraftId           = $Reader.GetValue(2)
      $FlightId             = $Reader.GetValue(3)
      $DepartureAirportIATA = $Reader.GetValue(4)
      $ArrivalAirportIATA   = $Reader.GetValue(5)
      $alertStatus = "N";
      $myString = "Cyberjet: <font color=#008000><b>FMC-PWI Msgs</b></font>" + " " + $LogtTS + " " + $MessageType + " " + $AircraftId + " " + $FlightId + " " + $DepartureAirportIATA + " " + $ArrivalAirportIATA 
      [void]$statusMsg.AppendLine($myString);
    } 
  } else {
      $myString = "Cyberjet: <font color=#F70D1A><b>No PWI Msgs </b></font>"
      [void]$statusMsg.AppendLine($myString);
    }
 
  $SQLcommand.Dispose();
  $Reader.Dispose();
  $Reader.Close();

  # * --------------------------------------------------------------------------------------
  # * Cyberjet: M40
  # * --------------------------------------------------------------------------------------
  [string] $SqlQuery = "SELECT TOP (1) [LogtTS],[MessageType],[AircraftId]," + 
                                      "[FlightId],[DepartureAirportIATA],"+
                                      "[ArrivalAirportIATA] " +
                                      "FROM [AircomSrv].[dbo].[Logt]  " +
                                      "WHERE [LogtTS] > DATEADD(Minute, -5, GETDATE())  " +
                                      " AND [Direction] = 2 " + 
                                      " AND [MessageType] LIKE '%M40 Messages From Cyberjet%' " + 
                                      "ORDER BY [LogtTS] DESC";

  $SQLcommand = New-Object System.Data.SqlClient.SqlCommand($SQLquery, $SQLconnection)
  $Reader = $SQLcommand.ExecuteReader()
  if ($Reader.HasRows) {
    while ($Reader.Read()) {
      $LogtTS               = $Reader.GetValue(0)
      $MessageType          = $Reader.GetValue(1)
      $AircraftId           = $Reader.GetValue(2)
      $FlightId             = $Reader.GetValue(3)
      $DepartureAirportIATA = $Reader.GetValue(4)
      $ArrivalAirportIATA   = $Reader.GetValue(5)
      $alertStatus = "N";
      $myString = "Cyberjet: <font color=#008000><b>M40 Messages</b></font>" + " " + $LogtTS + " " + $MessageType + " " + $AircraftId + " " + $FlightId + " " + $DepartureAirportIATA + " " + $ArrivalAirportIATA 
      [void]$statusMsg.AppendLine($myString);
    } 
  } else {
      $myString = "Cyberjet: <font color=#F70D1A><b>No M40 Msgs </b></font>"
      [void]$statusMsg.AppendLine($myString);
    }
 
  $SQLcommand.Dispose();
  $Reader.Dispose();
  $Reader.Close();

  # * --------------------------------------------------------------------------------------
  # * Cyberjet: PDC
  # * --------------------------------------------------------------------------------------
  [string] $SqlQuery = "SELECT TOP (1) [LogtTS],[MessageType],[AircraftId]," + 
                                      "[FlightId],[DepartureAirportIATA],"+
                                      "[ArrivalAirportIATA] " +
                                      "FROM [AircomSrv].[dbo].[Logt]  " +
                                      "WHERE [LogtTS] > DATEADD(Minute, -5, GETDATE())  " +
                                      " AND [Direction] = 2 " + 
                                      " AND ([MessageType] LIKE '%PDC Messages From Cyberjet%' " + 
                                      "   or [MessageType] LIKE '%PDC Messages From Cyberjet with Warning%') " + 
                                      "ORDER BY [LogtTS] DESC";

  $SQLcommand = New-Object System.Data.SqlClient.SqlCommand($SQLquery, $SQLconnection)
  $Reader = $SQLcommand.ExecuteReader()
  if ($Reader.HasRows) {
    while ($Reader.Read()) {
      $LogtTS               = $Reader.GetValue(0)
      $MessageType          = $Reader.GetValue(1)
      $AircraftId           = $Reader.GetValue(2)
      $FlightId             = $Reader.GetValue(3)
      $DepartureAirportIATA = $Reader.GetValue(4)
      $ArrivalAirportIATA   = $Reader.GetValue(5)
      $alertStatus = "N";
      $myString = "Cyberjet: <font color=#008000><b>PDC Messages</b></font>" + " " + $LogtTS + " " + $MessageType + " " + $AircraftId + " " + $FlightId + " " + $DepartureAirportIATA + " " + $ArrivalAirportIATA 
      [void]$statusMsg.AppendLine($myString);
    } 
  } else {
      $myString = "Cyberjet: <font color=#F70D1A><b>No PDC Msgs </b></font>"
      [void]$statusMsg.AppendLine($myString);
    }
 
  $SQLcommand.Dispose();
  $Reader.Dispose();
  $Reader.Close();

  # * --------------------------------------------------------------------------------------
  # * Cyberjet: NO PDC
  # * --------------------------------------------------------------------------------------
  [string] $SqlQuery = "SELECT TOP (1) [LogtTS],[MessageType],[AircraftId]," + 
                                      "[FlightId],[DepartureAirportIATA],"+
                                      "[ArrivalAirportIATA] " +
                                      "FROM [AircomSrv].[dbo].[Logt]  " +
                                      "WHERE [LogtTS] > DATEADD(Minute, -5, GETDATE())  " +
                                      " AND [Direction] = 2 " + 
                                      " AND [MessageType] LIKE '%PDC No Dep Clearance Message From Cyberjet%' " + 
                                      "ORDER BY [LogtTS] DESC";

  $SQLcommand = New-Object System.Data.SqlClient.SqlCommand($SQLquery, $SQLconnection)
  $Reader = $SQLcommand.ExecuteReader()
  if ($Reader.HasRows) {
    while ($Reader.Read()) {
      $LogtTS               = $Reader.GetValue(0)
      $MessageType          = $Reader.GetValue(1)
      $AircraftId           = $Reader.GetValue(2)
      $FlightId             = $Reader.GetValue(3)
      $DepartureAirportIATA = $Reader.GetValue(4)
      $ArrivalAirportIATA   = $Reader.GetValue(5)
      $alertStatus = "N";
      $myString = "Cyberjet: <font color=#F70D1A><b>PDC No PDCs </b></font>" + " " + $LogtTS + " " + $MessageType + " " + $AircraftId + " " + $FlightId + " " + $DepartureAirportIATA + " " + $ArrivalAirportIATA 
      [void]$statusMsg.AppendLine($myString);
    } 
  } else {
      $myString = "Cyberjet: <font color=#008000><b>No No PDCs  </b></font>No 'No Dep Clearance Message From Cyberjet' msg has been recevied"
      [void]$statusMsg.AppendLine($myString);
    }
 
  $SQLcommand.Dispose();
  $Reader.Dispose();
  $Reader.Close();

  # * --------------------------------------------------------------------------------------
  # * Cyberjet: Reject SOAR
  # * --------------------------------------------------------------------------------------
  [string] $SqlQuery = "SELECT TOP (1) [LogtTS],[MessageType],[AircraftId]," + 
                                      "[FlightId],[DepartureAirportIATA],"+
                                      "[ArrivalAirportIATA] " +
                                      "FROM [AircomSrv].[dbo].[Logt]  " +
                                      "WHERE [LogtTS] > DATEADD(Minute, -5, GETDATE())  " +
                                      " AND [Direction] = 2 " + 
                                      " AND [MessageType] LIKE '%FMC-Reject Messages From SOAR ACARS Adapter%' " + 
                                      "ORDER BY [LogtTS] DESC";

  $SQLcommand = New-Object System.Data.SqlClient.SqlCommand($SQLquery, $SQLconnection)
  $Reader = $SQLcommand.ExecuteReader()
  if ($Reader.HasRows) {
    while ($Reader.Read()) {
      $LogtTS               = $Reader.GetValue(0)
      $MessageType          = $Reader.GetValue(1)
      $AircraftId           = $Reader.GetValue(2)
      $FlightId             = $Reader.GetValue(3)
      $DepartureAirportIATA = $Reader.GetValue(4)
      $ArrivalAirportIATA   = $Reader.GetValue(5)
      $alertStatus = "N";
      $myString = "Cyberjet: <font color=#F70D1A><b>REJ SOAR ADP</b></font>" + " " + $LogtTS + " " + $MessageType + " " + $AircraftId + " " + $FlightId + " " + $DepartureAirportIATA + " " + $ArrivalAirportIATA 
      [void]$statusMsg.AppendLine($myString);
    } 
  } else {
      $myString = "Cyberjet: <font color=#008000><b>No SOAR Rej </b></font>No 'FMC-Reject Messages From SOAR ACARS Adapter'msg has been recevied"
      [void]$statusMsg.AppendLine($myString);
    }
 
  $SQLcommand.Dispose();
  $Reader.Dispose();
  $Reader.Close();

  # * --------------------------------------------------------------------------------------
  # * Database:  POS, FTM ACK, Cyberjet
  # * --------------------------------------------------------------------------------------
  [string] $SqlQuery = "SELECT TOP (1) [LogtTS],[MessageType],[AircraftId]," + 
                                      "[FlightId],[DepartureAirportIATA],"+
                                      "[ArrivalAirportIATA] " +
                                      "FROM [AircomSrv].[dbo].[Logt]  " +
                                      "WHERE [LogtTS] > DATEADD(Minute, -5, GETDATE())  " +
                                      " AND [Direction] = 4 " + 
                                      " AND ([MessageType] LIKE '%DB POS%' " + 
                                      "   or [MessageType] LIKE '%DB FTM ACK%' " + 
                                      "   or [MessageType] LIKE '%DB Cyberjet%') " + 
                                      "ORDER BY [LogtTS] DESC";

  $SQLcommand = New-Object System.Data.SqlClient.SqlCommand($SQLquery, $SQLconnection)
  $Reader = $SQLcommand.ExecuteReader()
  if ($Reader.HasRows) {
    while ($Reader.Read()) {
      $LogtTS               = $Reader.GetValue(0)
      $MessageType          = $Reader.GetValue(1)
      $AircraftId           = $Reader.GetValue(2)
      $FlightId             = $Reader.GetValue(3)
      $DepartureAirportIATA = $Reader.GetValue(4)
      $ArrivalAirportIATA   = $Reader.GetValue(5)
      $alertStatus = "N";
      $myString = "Database: <font color=#008000><b>DB Messages </b></font>" + $LogtTS + " " + $MessageType + " " + $AircraftId + " " + $FlightId + " " + $DepartureAirportIATA + " " + $ArrivalAirportIATA 
      [void]$statusMsg.AppendLine($myString);
    } 
  } else {
      $myString = "Database: <font color=#F70D1A><b>No DB  Msgs </b></font>"
      [void]$statusMsg.AppendLine($myString);
    }
 
  # * --------------------------------------------------------------------------------------
  # * Close and dispose of the connection and SQL objects
  # * Close out message
  # * Return the previously saved SQL command completion code
  # * --------------------------------------------------------------------------------------
  $SQLconnection.Close();
  $SQLconnection.dispose();
  $SQLcommand.dispose();

  [void]$statusMsg.AppendLine(" ");
  return $alertStatus
}

# * --------------------------------------------------------------------------------------
# * EXCEPTION PROCESSING
# * If any of the SQL objects were created, dispose of them and return not successful -1
# * --------------------------------------------------------------------------------------
catch {
  if ($SQLconnection) {
      $SQLconnection.Close();
      $SQLconnection.dispose();
  }
  if ($SQLcommand) {
      $SQLcommand.dispose();
  }
  $myErrorMsg = "ERROR* * * in checkCyberjetStatus " + " " + $myVersion + " $($_.Exception.Message)";
  [void]$statusMsg.AppendLine("");
  [void]$statusMsg.AppendLine($myErrorMsg);
  [void]$statusMsg.AppendLine("" + "</PRE></font>");

  $sendMailMessageSplat = @{
    From = $from
    To = "<John.Buxton@aa.com>"
    Subject = "ASP Status Report ** Exception ** "
    SmtpServer = 'smtp.qcorpaa.aa.com'
    Body=$myErrorMsg
  }
  Send-MailMessage @sendMailMessageSplat

  return "Y"
}

