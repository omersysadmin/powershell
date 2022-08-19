
<#
.SYNOPSIS
Script to Gather the necessairy information for computer inventory
.DESCRIPTION
This is a script I use for computer inventory, only thing that is still not working is the moniotor information gathering.
#>


function Get-InfoMachine {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $ComputerName
    )

    $output = @{
        
        Marque = ""
        Adresse_IP = ""
        Numero_de_serie = ""

    }
    
    if (Test-Connection -ComputerName $ComputerName -Count 1 -ErrorAction SilentlyContinue){
        Write-Debug "$ComputerName connectée"
        $output.Numero_de_serie = (Get-WmiObject -Class win32_bios -ComputerName $ComputerName).serialnumber
        # $Desktop =  Get-WmiObject -Class win32_computersystem -ComputerName $ComputerName
        $output.Marque = (get-wmiobject -Class win32_computersystem -ComputerName $ComputerName).Manufacturer
        $output.Model = (get-wmiobject -Class win32_computersystem -ComputerName $ComputerName).systemfamily
        $network = get-wmiobject -class win32_networkadapterconfiguration | where-Object Ipenabled -eq $true
        $output.Adresse_IP = $network.IPAddress[0]
        $output.MACAddress = $network.MACAddress
        $output.System = (Get-WmiObject -Class win32_operatingsystem).Caption
        $output.Office = (Get-WmiObject -Class win32_product -ComputerName $ComputerName | Where-Object {$_.Name -like "Microsoft Office Standard*" -or $_.Name -like "Microsoft Office Professional*"}).Name
        
        

        [PSCustomObject]$output
    }else {
        Write-Host "$ComputerName non connectée"
    }

}
   

Get-InfoMachine localhost