# Price War - automatic runs (made by SCHEDULE DAILY RUN.bat)
#   07:30 every day         : full check - 80% revenue SKUs on marketplaces + all QC apps
#   every 2 hours, all day  : Top 50 revenue SKUs + all QC apps (09:30, 11:30 ... 05:30)
# Runs hidden, also on battery, wakes the laptop from sleep, and catches up a missed run.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$vbs  = Join-Path $root "run_hidden.vbs"

foreach ($old in "QC Price Tracker daily", "Price War - daily 80pc", "Price War - Top 50 every 2 hours") {
    Unregister-ScheduledTask -TaskName $old -Confirm:$false -ErrorAction SilentlyContinue
}

$set = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
       -StartWhenAvailable -WakeToRun -MultipleInstances IgnoreNew `
       -ExecutionTimeLimit (New-TimeSpan -Minutes 100)

function Add-Run($name, $tier, $times) {
    $action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument ('"' + $vbs + '" --tier ' + $tier) -WorkingDirectory $root
    $trig = $times | ForEach-Object { New-ScheduledTaskTrigger -Daily -At $_ }
    Register-ScheduledTask -TaskName $name -Action $action -Trigger $trig -Settings $set `
        -Description "Price War (Nat Habit price parity) - $root" -Force | Out-Null
    $i = Get-ScheduledTask -TaskName $name | Get-ScheduledTaskInfo
    Write-Host ("  OK  {0,-36} next run: {1}" -f $name, $i.NextRunTime)
}

$d = [datetime]::Today
Add-Run "Price War - daily 80pc" "80" @($d.AddHours(7.5))
Add-Run "Price War - Top 50 every 2 hours" "top50" ((0..10) | ForEach-Object { $d.AddHours(((9 + 2 * $_) % 24) + 0.5) })
