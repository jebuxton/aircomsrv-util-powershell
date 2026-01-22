ToggleSQLAgent -Operation Enable

function RunSQL{
    param(
        [Parameter(Position=0)]
        [string]$Query,
        [Parameter(Position=1)]
        [string]$ServerInstance = $null
    )
     if ($PSVersionTable.PSVersion.Major -ge 6){
        if ($ServerInstance){
            Invoke-Sqlcmd -TrustServerCertificate -Query $Query -ServerInstance $ServerInstance
        } else {
            Invoke-Sqlcmd -TrustServerCertificate -Query $Query
        }
    }
    else{
        if ($ServerInstance){
            Invoke-Sqlcmd -Query $Query -ServerInstance $ServerInstance
        } else {
            Invoke-Sqlcmd -Query $Query
        }
    }
}


function ToggleSQLAgent {

# **********************************************************

# * Disable SQL SERVER Agents on Inactive datacenter

# * 1) Monitor: Aircom-Archive-Job

# * 2) POS Report Monitor

# **********************************************************

param(

    [Parameter(Mandatory=$true)]

    [ValidateSet('Disable','Enable')]

    [string]$Operation

)

  $jobsToToggle = @(

    "Monitor: Aircom-Archive-Job",

    "POS Report Monitor"

  )
 
  $enabled = 0

  if ($Operation -eq 'Enable'){ $enabled = 1 }

  else { $enabled = 0 }
 
  foreach ($jobName in $jobsToToggle) {

    $sql = @"

      EXEC msdb.dbo.sp_update_job 

        @job_name = N'$jobName',

        @enabled = $enabled;

"@

    RunSQL -Query $sql

    if ($Operation -eq 'Enable'){

      Write-Host "Job '$jobName' has been enabled."

    }

    else{

      Write-Host "Job '$jobName' has been disabled."

    }

  }

}
 
