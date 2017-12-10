
##
###########################################################################################################
##                                        Common Functions                                               ##
###########################################################################################################
##


Function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [switch]$Warning = $false
    )
    $Timestamp = Get-Date -Format "MM-dd-yyyy hh:mm:ss"
    Write-Host -NoNewline -ForegroundColor White "[$Timestamp]"

    if($Warning){
        Write-Host -ForegroundColor Yellow " WARNING: $Message"
    } else {
        Write-Host -ForegroundColor Green " $Message"
    }
    $LogMessage = "[$Timestamp] $Message"
    $LogMessage | Out-File -Append -LiteralPath $VerboseLogFile
}


function Get-VCSAConnection {
    param(
        [string]$vcsaName,
        [string]$vcsaUser,
        [string]$vcsaPassword
    )
    $existingConnection =  $global:DefaultVIServers | where-object -Property Name -eq -Value $vcsaName
    if($existingConnection -ne $null) {
        return $existingConnection;
    } else {
        $connection = Connect-VIServer -Server $vcsaName -User $vcsaUser -Password $vcsaPassword -WarningAction SilentlyContinue;
        return $connection;
    }
}

function Close-VCSAConnection {
    param(
        [string]$vcsaName
    )
    if($vcsaName.Length -le 0) {
        Disconnect-VIServer -Server $Global:DefaultVIServers -Confirm:$false
    } else {
        $existingConnection =  $global:DefaultVIServers | where-object -Property Name -eq -Value $vcsaName
        if($existingConnection -ne $null) {
            Disconnect-VIServer -Server $existingConnection -Confirm:$false;
        } else {
            Write-Log -Message "Could not find an existing connection named $($vcsaName)" -Warning
        }
    }
}
