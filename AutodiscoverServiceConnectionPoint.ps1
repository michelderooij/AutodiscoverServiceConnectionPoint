<#
    .SYNOPSIS
    AutodiscoverServiceConnectionPoint.ps1
   
    Michel de Rooij
    michel@eightwone.com
	 
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.0, September 30th, 2015
    
    .DESCRIPTION
    This script contains functions to create, remove and configure 
    Autodiscover Service Connection Points for Exchange.
	
    .LINK
    http://eightwone.com
    
    .NOTES
    Requirements:
    - AD PowerShell Module

    Revision History
    --------------------------------------------------------------------------------
    1.0     Initial community release

    .EXAMPLE
    Clear-AutodiscoverServiceConnectionPoint -Name EX1

    .EXAMPLE
    Clear-AutodiscoverServiceConnectionPoint -Name EX2

    .EXAMPLE
    New-AutodiscoverServiceConnectionPoint -Name EX1 -ServiceBinding https://autodiscover.contoso.com/autodiscover/autodiscover.xml -Site Default-First-Site-Name

    .EXAMPLE
    New-AutodiscoverServiceConnectionPoint -Name EX2 -ServiceBinding https://autodiscover.contoso.com/autodiscover/autodiscover.xml

    .EXAMPLE
    Set-AutodiscoverServiceConnectionPoint -Name EX2 -Site Default-First-Site-Name
#>

    Function Get-ADConfRoot() {
        Return 'cn=Configuration,' + (Get-ADDomain).DistinguishedName
    }

    Function Get-AutodiscoverServiceConnectionPoint ([string]$Name) {
        $ConfRoot= Get-ADConfRoot
        Write-Verbose "Looking for Exchange-AutoDiscover-Service record for $Name"
        $SCP= Get-ADObject -SearchBase $ConfRoot -LDAPFilter "(&(cn=$($Name))(objectClass=serviceConnectionPoint)(serviceClassName=ms-Exchange-AutoDiscover-Service)(|(keywords=67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68)(keywords=77378F46-2C66-4aa9-A6A6-3E7A48B19596)))" -Properties *
        Return $SCP
    }

    Function Clear-AutodiscoverServiceConnectionPoint ([string]$Name) {
        $SCP= Get-AutodiscoverServiceConnectionPoint $Name
        Write-Verbose "Removing $($SCP.distinguishedName)"
        $SCP | Remove-ADObject -Confirm:$false
    }

    Function Get-OrganizationName () {
        $ConfRoot= Get-ADConfRoot
        $OrgObject= Get-ADObject -LDAPFilter '(objectClass=msExchOrganizationContainer)' -SearchBase $ConfRoot 
        If( $OrgObject.Count -gt 1) {
            Write-Error 'Multiple Exchange Organization objects found in AD, aborting'
            Exit 1
        }
        Else {
            Write-Verbose "Exchange Organization Name is $($OrgObject.Name)"
        }
        Return ($OrgObject).Name
    }

    Function Set-AutodiscoverServiceConnectionPoint( [string]$Name, [string]$ServiceBinding, [string]$Site) {
        $SCP= Get-AutodiscoverServiceConnectionPoint $Name
        If( $SCP) {
            If( $ServiceBinding) { $SCP.ServiceBinding= $ServiceBinding }
            If( $Site) { [array]$SCP.Keywords+= @('77378F46-2C66-4aa9-A6A6-3E7A48B19596', "Site=$Site") }
        }
        Else {
            Write-Warning 'No Service Connection Point for $Name found'
        }
    }

    Function New-AutodiscoverServiceConnectionPoint( [string]$Name, [string]$ServiceBinding, [string]$Site) {
        If( Get-AutodiscoverServiceConnectionPoint $Name) {
            Write-Warning 'Service Connection Point for $Name already exists'
        }
        Else {
            $ConfRoot= Get-ADConfRoot
            $OrgName= Get-OrganizationName
            $ProjectedPath= "CN=Autodiscover,CN=Protocols,CN=$Name,CN=Servers,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN=$OrgName,CN=Microsoft Exchange,CN=Services,$ConfRoot"
            $Keywords= @('77378F46-2C66-4aa9-A6A6-3E7A48B19596')
            If( $Site) {
                $Keywords+= "Site=$Site"
            }
            Write-Host "Creating ms-Exchange-AutoDiscover-Service record for $Name at $ProjectedPath"
            New-ADObject -Name $Name -Type serviceConnectionPoint -Path $ProjectedPath -OtherAttributes @{
                Keywords=$Keywords; 
                serviceBindingInformation=$ServiceBinding; 
                serviceClassName="ms-Exchange-AutoDiscover-Service";
                serviceDNSName="$Name" }
        }
    }

