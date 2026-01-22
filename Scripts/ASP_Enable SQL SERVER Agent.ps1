EnableSQLAgent -ServerName "PSSAPPT6984\AIRCOM_DB"

function EnableSQLAgent {
# **********************************************************
# * Enable SQL SERVER Agents on Active datacenter
# * 1) Monitor: Aircom-Archive-Job
# * 2) POS Report Monitor
# **********************************************************
    param (
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )

    $jobsToEnable = @(
        "Monitor: Aircom-Archive-Job",
        "POS Report Monitor"
    )

    foreach ($jobName in $jobsToEnable) {
        try {
            $sql = "EXEC msdb.dbo.sp_update_job @job_name = N'$jobName', @enabled = 1;"
            Invoke-Sqlcmd -ServerInstance $ServerName -Query $sql -ErrorAction Stop
            Write-Host "Job '$jobName' has been enabled on server '$ServerName'."
        }
        catch {
            Write-Warning "Failed to enable job '$jobName' on server '$ServerName'. Error: $_"
        }
    }
}
