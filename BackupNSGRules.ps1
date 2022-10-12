param (
    [Parameter(Mandatory = $false,ParameterSetName="All")][Switch]$All,
    [Parameter(Mandatory = $false,ParameterSetName="Subscription")][string]$SubscriptionName
)

function ConvertTo-ConsoleSafeString {
    param (
        [Parameter(Mandatory = $true)][string]$String
    )
    if ($String -ne $null) {
        # Bad regex. Make better if time permits.
        return "$("$("$($String -replace '\r', '`r')" -replace '\n', '`n')" -replace '\t', ' ')".Replace('"','""')
    }
    return
}

function ConvertTo-StringArray {
    param (
        [Parameter(Mandatory = $true)][string]$String
    )
    foreach ($element in $String.Split(" ")) {
        if ($Result) {$Result = ($Result, "," -join "")}
        $Result = $($Result, """$($element)""" -join "")

    }
    return "($($Result))"
}




try {
    $Context = Get-AzContext
}
catch [CouldNotAutoloadMatchingModule] {
    Write-Host "Could not load the appropriate module. Run `"Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted`" to proceed."
    exit
}

#Set Variables
$Location = "$($Context.Environment.Name)"
$Date = Get-Date
$DirectoryTimestamp=$(Get-Date -Date $Date -UFormat "%Y%m - %B")
$FileTimestamp=$(Get-Date -Date $Date -Format "yyyyMMddHHmmss")

$UsageSummaryPath=$(Convert-Path .)
$SummaryFileName = "NSG-Summary-$($Location)-$($FileTimestamp).csv"

$OutputDirectory = "$($UsageSummaryPath)\NetworkSecurityGroupRules\$($DirectoryTimestamp)"


if (-not (Test-Path "$OutputDirectory")) {
    New-Item -ItemType Directory -Path "$OutputDirectory" | Out-Null}

#Get Subscription
if ($SubscriptionName) {
    $TenantId = (Get-AzContext).Tenant.Id
    $Subscriptions = Get-AzSubscription -SubscriptionName $SubscriptionName -TenantId $TenantId
}
else {
    
    $Subscriptions = Get-AzSubscription
}

#Display Progress
Write-Host "Writing to file $($OutputDirectory)\$($SummaryFileName)"
$Summary = @()
$SubscriptionProgress = 0
$SubscriptionCount = $($Subscriptions).Count

#Create README
New-Item -Path $($OutputDirectory) -Name README.txt -Force
Set-Content "$($OutputDirectory)\README.txt" "To deploy a template,`r`n 1. Copy the template file to your PowerShell working directory. `r`n 2. Update 'Recource_Group_Name' and 'Template_file_name.json' in the command below and run it in PowerShell. `r`n New-AzResourceGroupDeployment -ResourceGroupName 'Recource_Group_Name' -TemplateFile 'Template_file_name'"

#Export NSG rules
foreach ($Subscription in $Subscriptions) {
    $SubscriptionProgress += 1
    Write-Progress -Id 0 -CurrentOperation "$($Subscription.Name)" -Status "$([Int]($SubscriptionProgress/$SubscriptionCount*100))% Complete:" -PercentComplete ($SubscriptionProgress/$SubscriptionCount*100) -Activity "Iterating through subscriptions."

    $NetworkSecurityGroups = Get-AzNetworkSecurityGroup



    $NetworkSecurityGroupProgress = 0
    $NetworkSecurityGroupCount = $($NetworkSecurityGroups).Count

    foreach ($NetworkSecurityGroup in $NetworkSecurityGroups) {
        
        $OutputJsonName = "$($NetworkSecurityGroup.Name)-$($FileTimestamp).json"

        Export-AzResourceGroup -ResourceGroupName $($NetworkSecurityGroup.ResourceGroupName) -SkipAllParameterization -Resource $($NetworkSecurityGroup.Id) -Path $($OutputDirectory) | Rename-Item -NewName $OutputJsonName
        

        $NetworkSecurityGroupProgress += 1
        Write-Progress -Id 1 -ParentId 0 -CurrentOperation "$($NetworkSecurityGroup.Name)" -Status "$([Int]($NetworkSecurityGroupProgress/$NetworkSecurityGroupCount*100))% Complete:" -PercentComplete ($NetworkSecurityGroupProgress/$NetworkSecurityGroupCount*100) -Activity "Iterating through network security groups."

       

        $SecurityRuleProgress = 0
        $SecurityRuleCount = $NetworkSecurityGroup.SecurityRules.Count
        foreach ($SecurityRule in $NetworkSecurityGroup.SecurityRules) {
            $Record = New-Object -TypeName System.Object
    
            $Record | Add-Member -Name "Subscription" -MemberType NoteProperty -Value "$($Subscription.Name)"
            $Record | Add-Member -Name "NetworkSecurityGroup" -MemberType NoteProperty -Value "$($NetworkSecurityGroup.Name)"
            $Record | Add-Member -Name "Location" -MemberType NoteProperty -Value "$($NetworkSecurityGroup.Location)"
            $Record | Add-Member -Name "ApplicableSubnet" -MemberType NoteProperty -Value "$($SubnetName)".Replace(" ", ", ")
            $Record | Add-Member -Name "ApplicableInterfaces" -MemberType NoteProperty -Value "$($NetworkInterfaceName)".Replace(" ", ", ")
            $Record | Add-Member -Name "Direction" -MemberType NoteProperty -Value "$($SecurityRule.Direction)"
            $Record | Add-Member -Name "Priority" -MemberType NoteProperty -Value "$($SecurityRule.Priority)"
            $Record | Add-Member -Name "Access" -MemberType NoteProperty -Value "$($SecurityRule.Access)"
            $Record | Add-Member -Name "Name" -MemberType NoteProperty -Value "$($SecurityRule.Name)"
            $Record | Add-Member -Name "Source" -MemberType NoteProperty -Value "$($SecurityRule.SourceAddressPrefix)".Replace(" ", ", ")
            $Record | Add-Member -Name "SourcePort" -MemberType NoteProperty -Value "$($SecurityRule.SourcePortRange)".Replace(" ", ", ")
            $Record | Add-Member -Name "Destination" -MemberType NoteProperty -Value "$($SecurityRule.DestinationAddressPrefix)".Replace(" ", ", ")
            $Record | Add-Member -Name "DestinationPort" -MemberType NoteProperty -Value "$($SecurityRule.DestinationPortRange)".Replace(" ", ", ")
            $Record | Add-Member -Name "Protocol" -MemberType NoteProperty -Value "$($SecurityRule.Protocol)"
            $Record | Add-Member -Name "Description" -MemberType NoteProperty -Value "$($SecurityRule.Description)"
            $Record | Export-Csv -Path "$($OutputDirectory)\$($SummaryFileName)" -Append -NoTypeInformation

           
          
        } 
        
        
    }
}
