<#
.Synopsis
   Distributed Cache Utilities
.DESCRIPTION
   Collection of Advanced Functions for working with Distributed Cache
   Depends on PS Remoting available on every server in the farm
.NOTES
	File Name  : SharePointDistributedCache.psm1
	Author     : Spencer Harbar (spence@harbar.net)
	Requires   : PowerShell Version 4.0  
#>



Function Update-DistributedCacheIdentity {
<#
.Synopsis
   Updates the Distributed Cache Service Account Identity.
.DESCRIPTION
   Updates the Distributed Cache Service Account Identity.
   Can be run from any Server in the Farm.
   If the Farm Service exists and a Cache Host exists, it will deploy after ensuring there is only
   one Cache Host and then re-add the other Cache Hosts
   If the Farm Service exists but zero Cache Hosts exist, and we include -InitialIdentity it will update 
   the identity but not deploy, which supports a SharePoint 2016 MinRole Farm build pattern. Should never be 
   used afer creating the Cache Cluster

   If Deploy fails we put the Cluster back as it was
.EXAMPLE
   Update-DistributedCacheIdentity -ManagedAccountName "DOMAIN\user" -Credential $Credential
.EXAMPLE
   Update-DistributedCacheIdentity -ManagedAccountName "DOMAIN\user" -Credential $Credential -InitialIdentity
#>

    [CmdletBinding()]
    Param (  
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="The Name of an existing Managed Account to use as the service identity, e.g. DOMAIN\user",
                   Position=0)]
        [ValidateNotNullorEmpty()]
        [String]
        $ManagedAccountName,
        
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="The Credential of a Shell Admin",
                   Position=1)]
        [ValidateNotNullorEmpty()]
        [PSCredential]
        $Credential,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="For SharePoint 2016 Farms where no Distributed Cache machine has yet been added",
                   Position=2)]
        [Switch]$InitialIdentity
    ) 

    Begin {
        Write-Output "$(Get-Date -Format T) : Initiated Distributed Cache Service Identity change." 
        if ((Get-PSSnapin -Name "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) {
            Add-PSSnapin -Name "Microsoft.SharePoint.PowerShell"
        }     
        $ScriptBlock = {

            Write-Output "$(Get-Date -Format T) : Initiated Distributed Cache Service Identity deploy." 
            Add-PSSnapin -Name "Microsoft.SharePoint.PowerShell"

            try {
                $Farm = Get-SPFarm
                $DistributedCacheService = $Farm.Services | Where-Object { $_.TypeName -eq "Distributed Cache" }
                Write-Output "$(Get-Date -Format T) : This will take approximately 8 minutes..."
                $DistributedCacheService.ProcessIdentity.Deploy()
                Write-Output "$(Get-Date -Format T) : Distributed Cache Service Account Change Deployed!"

            } catch {
                Write-Error "ERROR: We failed during changing the Distributed Cache Service Account."
                $_.Exception.Message
            }
        }
    }

    Process {
        # Grab the Farm, Managed Account, Farm Service and the Servers running DC
        $Farm = Get-SPFarm
        $ManagedAccount = Get-SPManagedAccount -Identity $ManagedAccountName
        $DistributedCacheService = $Farm.Services | Where-Object { $_.TypeName -eq "Distributed Cache" }
        $DistributedCacheServers = Get-SPServer | 
                                   Where-Object {($_.ServiceInstances | ForEach TypeName) -eq "Distributed Cache"} | 
                                   ForEach Address

        # if the service exsits the farm, we can change the identity...
        # but only if no instance has ever been started, otherwise this cause "unkown" after host is brought up
        If ($DistributedCacheService) {

            # No servers running DC, we fall thru as there is no further work to do
            # unless InitialIdentity is specified, in which case we update but don't deploy
            If ($DistributedCacheServers.Count -eq 0) {
                If ($InitialIdentity) {
                    Write-Output "$(Get-Date -Format T) : Configuring Distributed Cache Service Account..." 
                    $DistributedCacheService.ProcessIdentity.CurrentIdentityType = "SpecificUser"
                    $DistributedCacheService.ProcessIdentity.ManagedAccount = $ManagedAccount
                    $DistributedCacheService.ProcessIdentity.Update() 
                    Write-Output "$(Get-Date -Format T) : Distributed Cache Service Account Updated!"    
                }
                Else {
                    Write-Output "$(Get-Date -Format T) : No servers in this farm are running Distributed Cache!"
                }
            }
            # there is at least one server
            Else {
                Write-Output "$(Get-Date -Format T) : Configuring Distributed Cache Service Account..." 
                $DistributedCacheService.ProcessIdentity.CurrentIdentityType = "SpecificUser"
                $DistributedCacheService.ProcessIdentity.ManagedAccount = $ManagedAccount
                $DistributedCacheService.ProcessIdentity.Update() 
                Write-Output "$(Get-Date -Format T) : Distributed Cache Service Account Updated!"                 
                # If More than one DC server, remove all but one
                For ($i=0; $i -le ($DistributedCacheServers.Count - 2); $i++) {
                    Remove-DistributedCache -ComputerName $DistributedCacheServers[$i] -Credential $Credential
                }
                        
                # Get the remaining server
                $ChangeIdentityServer = Get-SPServer | 
                                        Where-Object {($_.ServiceInstances | ForEach TypeName) -eq "Distributed Cache"} | 
                                        ForEach Address

                #run the deploy() script block
                Invoke-Command -ComputerName $ChangeIdentityServer -Credential $Credential -Authentication Credssp `
                               -ScriptBlock $ScriptBlock

                # add back the other DC servers
                ForEach ($DistributedCacheServer in $DistributedCacheServers) {
                    If ($DistributedCacheServer -ne $ChangeIdentityServer) {
                        Add-DistributedCache -ComputerName $DistributedCacheServer -Credential $Credential
                    }
                }
            } 
        }
        Else {
            Write-Output "$(Get-Date -Format T) : Distributed Cache Service does not yet exist in the Farm!"
        }
    }

    End {}
}

Function Update-DistributedCacheSize {
<#
.Synopsis
   Updates the Cache Size in MB on every Cache Host in the Cluster.
.DESCRIPTION
   Updates the Cache Size in MB on every Cache Host in the Cluster.
   Can be run from any server in the farm.
   Stops the Service Instance where ever it is running, configures the Cache Size and Starts the Service Instances

   SP2016 Note:
   Once we stop the SIs each machine will be out of compliance with minrole.
   All will be good once it's complete, however we could use Stop-SPService first which sets AutoProvision to false.
   Then Start-SPService after changing the cache size. that way there can never be any "conflict"
   But that only works with SP16 and MinRole. 
   This way it works no matter which version or if we are MinRole or Custom
   The time it takes is so small, and you shoyuld be aware of when timer jobs are running anyway!

.EXAMPLE
   Update-DistributedCacheSize -CacheSizeInMB 500 -Credential $Credential
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="The Cache Size in MB to configure on each Cache Host",
                   Position=0)]
        [ValidateNotNullorEmpty()]
        [Int]
        $CacheSizeInMB,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="The Credential of a Shell Admin",
                   Position=1)]
        [ValidateNotNullorEmpty()]
        [PSCredential]
        $Credential
    )

    Begin {
        Write-Output "$(Get-Date -Format T) : Initiated Distributed Cache Service Cache Size change."
        if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) {
            Add-PSSnapin "Microsoft.SharePoint.PowerShell"
        }
         $ScriptBlock = {
            param (
                [Int]$CacheSizeInMB
            )
            Add-PSSnapin -Name "Microsoft.SharePoint.PowerShell"
            Update-SPDistributedCacheSize -CacheSizeInMB $CacheSizeInMB
         }
    }

    Process {
        try {
            # Get the servers running DC
            $DistributedCacheServers = @( Get-SPServer | `
                                       Where-Object {($_.ServiceInstances | ForEach TypeName) -eq "Distributed Cache"} | `
                                       ForEach Address )

            # Stop service instances
            Stop-DistributedCacheCluster

            # Update the Cache Size
            Write-Output "$(Get-Date -Format T) : Changing Distributed Cache Service Cache Size..."
            Invoke-Command -ComputerName $DistributedCacheServers[0] -Credential $Credential -Authentication Credssp `
                           -ArgumentList $CacheSizeInMB -ScriptBlock $ScriptBlock

            # Start service instances
            Start-DistributedCacheCluster
        }
        catch {
            Write-Error "ERROR: We failed during changing the Distributed Cache Size."
            $_.Exception.Message
        }
    }

    End {
        Write-Output "$(Get-Date -Format T) : Completed Distributed Cache Service Cache Size Change!"
    }
}

Function Start-DistributedCacheCluster {
<#
.Synopsis
   Starts the Distributed Cache Service Instance on every server it exists
.DESCRIPTION
   Starts the Distributed Cache Service Instance on every server it exists.
   Can be run from any server in the farm.
   Used in the 'update' pattern, for example changing the Cache Size.
   It is expected that a previous Cache Cluster has been created using 
   Add-SPDistributedCacheServiceInstance or the Server Role Distributed Cache
.EXAMPLE
   Start-DistributedCacheCluster
#>

    [CmdletBinding()]
    Param ()

    Begin {
        if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) {
            Add-PSSnapin "Microsoft.SharePoint.PowerShell"
        }
    }

    Process {
        Write-Output "$(Get-Date -Format T) : Starting Distributed Cache Service Instance on all servers..."
        Get-SPServiceInstance | Where-Object { $_.TypeName -eq "Distributed Cache" } | Start-SPServiceInstance | Out-Null
        While (Get-SPServiceInstance | Where-Object { $_.TypeName -eq "Distributed Cache" -and $_.Status -ne "Online" }) {
            Start-Sleep -Seconds 15
        }
    }

    End {
        Write-Output "$(Get-Date -Format T) : All Distributed Cache Service Instances started!"
    }
}

Function Stop-DistributedCacheCluster {
<#
.Synopsis
   Stops the Distributed Cache Service Instance on every server it exists
.DESCRIPTION
   Stops the Distributed Cache Service Instance on every server it exists.
   Can be run from any server in the farm.
   Used in the 'update' pattern, for example changing the Cache Size.
   It is expected that a previous Cache Cluster has been created using 
   Add-SPDistributedCacheServiceInstance or the Server Role Distributed Cache
.EXAMPLE
   Stop-DistributedCacheCluster
#>

    [CmdletBinding()]
    Param ()

    Begin {
        if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) {
            Add-PSSnapin "Microsoft.SharePoint.PowerShell"
        }
    }

    Process {
        Write-Output "$(Get-Date -Format T) : Stopping Distributed Cache Service Instance on all servers..."
        Get-SPServiceInstance | Where-Object { $_.TypeName -eq "Distributed Cache" } | Stop-SPServiceInstance -Confirm:$false | Out-Null
        While (Get-SPServiceInstance | Where-Object { $_.TypeName -eq "Distributed Cache" -and $_.Status -ne "Disabled" }) {
            Start-Sleep -Seconds 15
        }
    }

    End {
        Write-Output "$(Get-Date -Format T) : All Distributed Cache Service Instances stopped!"
    }
}

Function Get-DistributedCacheStatus {
<#
.Synopsis
   Displays the status of the Cache Cluster and optionally each Cache Host
.DESCRIPTION
   Displays the status of the Cache Cluster and optionally each Cache Host.
   Can be run from any server in the farm.
   Status of each host is the key - Up is All Good. Anything else means a broken cache cluster.
   Pass in -CacheHostConfiguration to view additional settings for each Cache Host (e.g. Cache Size).

   # Note (the bool should be a switch - won't flow thru as param)

.EXAMPLE
   Get-DistributedCacheStatus -Credential $Credential
.EXAMPLE
   Get-DistributedCacheStatus -Credential $Credential -CacheHostConfiguration
#>

    [CmdletBinding()]
    Param (  
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="The Credential of a Shell Admin",
                   Position=0)]
        [ValidateNotNullorEmpty()]
        [PSCredential]
        $Credential,

        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Output individual Cache Host Configuration",
                   Position=1)]
        [Switch]
        $CacheHostConfiguration
    ) 

    Begin {
        if ((Get-PSSnapin -Name "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) {
            Add-PSSnapin -Name "Microsoft.SharePoint.PowerShell"
        }
        $ScriptBlock = {

            param (
                [String[]]$DistributedCacheServers, 
                [Bool]$CacheHostConfiguration = $false
            )
            
            Use-CacheCluster
            #status of each host, first hit always seems to error on one of them!
            $DistributedCacheServers | ForEach {Get-AFCacheHostStatus -ComputerName $_ -CachePort 22233} | Format-Table

            if ($CacheHostConfiguration) {
                #config of each host
                $DistributedCacheServers | ForEach {Get-AFCacheHostConfiguration -ComputerName $_ -CachePort 22233} | 
                                           Select HostName, Size, HighWatermark, LowWatermark, IsLeadHost | Format-Table
            }

        }
        
    }

    Process {

        $DistributedCacheServers = @(Get-SPServer | 
                                   Where-Object {($_.ServiceInstances | ForEach TypeName) -eq "Distributed Cache"} | 
                                   ForEach Address)

        Invoke-Command -ComputerName $DistributedCacheServers[0] -Credential $Credential -Authentication Credssp `
                       -ArgumentList $DistributedCacheServers, $CacheHostConfiguration -ScriptBlock $ScriptBlock
    }

    End {}
}

Function Remove-DistributedCache {
<#
.Synopsis
   Removes a Cache Host from the Cluster and the Service Instance.
.DESCRIPTION
   Removes a Cache Host from the Cluster and the Service Instance.
   Can be run from any server in the farm.
   Waits for Port to free up
   Used in the 'identity' pattern which requires only a single Cache Host.
   Also useful when working more generally with Distributed Cache
.EXAMPLE
   Remove-DistributedCache -ComputerName "FABSP01" -Credential $Credential
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="The Computer Name of a machine to remove as a  Cache Host/Service Instance",
                   Position=0)]
        [ValidateNotNullorEmpty()]
        [String]
        $ComputerName,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="The Credential of a Shell Admin",
                   Position=1)]
        [ValidateNotNullorEmpty()]
        [PSCredential]
        $Credential
    ) 

    Begin {
        $ScriptBlock = {
            $server = $env:ComputerName
            Add-PSSnapin -Name "Microsoft.SharePoint.PowerShell"
            
            try {
                Write-Output "$(Get-Date -Format T) : Removing server $server as Distributed Cache host..."
                Remove-SPDistributedCacheServiceInstance -ErrorAction Stop
                $Count = 0
                $MaxCount = 5
                While ( ($Count -lt $MaxCount) -and (Get-NetTCPConnection -LocalPort 22234 -ErrorAction SilentlyContinue).Count -gt 0 ) {
                    Write-Output "$(Get-Date -Format T) : Waiting on port to free up..."
                    Start-Sleep -Seconds 30
                    $Count++
                }
                Write-Output "$(Get-Date -Format T) : Removed server $server as Distributed Cache host!"

            } catch {
                Write-Error "ERROR: We failed during removing cache host cluster on $server."
                $_.Exception.Message
            }
        }
    }

    Process {
        Invoke-Command -ComputerName $ComputerName -Credential $Credential -Authentication Credssp -ScriptBlock $ScriptBlock
    }

    End {}
}

Function Add-DistributedCache {
<#
.Synopsis
   Adds a Cache Host to the Cluster and the Service Instance.
.DESCRIPTION
   Adds a Cache Host to the Cluster and the Service Instance.
   Can be run from any server in the farm.
   Used in the 'identity' pattern which requires only a single Cache Host.
   Also useful when working more generally with Distributed Cache
.EXAMPLE
   Add-DistributedCache -ComputerName "FABSP01" -Credential $Credential
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="The Computer Name of a machine to add as a Cache Host/Service Instance",
                   Position=0)]
        [ValidateNotNullorEmpty()]
        [String]
        $ComputerName,

        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="The Credential of a Shell Admin",
                   Position=1)]
        [ValidateNotNullorEmpty()]
        [PSCredential]
        $Credential
    ) 

    Begin {
        $ScriptBlock = {
            $server = $env:ComputerName
            Add-PSSnapin -Name "Microsoft.SharePoint.PowerShell"
            
            try {
                Write-Output "$(Get-Date -Format T) : Adding server $server as Distributed Cache host..."
                Add-SPDistributedCacheServiceInstance -ErrorAction Stop
                Write-Output "$(Get-Date -Format T) : Added server $server as Distributed Cache host!"
            } catch {
                Write-Error "ERROR: We failed during adding cache host cluster on $server."
                $_.Exception.Message
            }
        }
    }

    Process {
        Invoke-Command -ComputerName $ComputerName -Credential $Credential -Authentication Credssp -ScriptBlock $ScriptBlock
    }

    End {}
}

#EOF