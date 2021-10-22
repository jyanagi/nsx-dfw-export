## Attribution Comments
# 
# Convert-OutputForCSV is created by Boe Prox (GITHUB proxb - https://github.com/proxb/PowerShell_Scripts/blob/master/Convert-OutputForCSV.ps1) 
# all credit goes to him for the Convert-OutputForCSV function)
# 
# Modified Convert-OutputForCSV for SemiColon functionality
#
## End Attribution Comments

Function Convert-OutputForCSV {
    <#
        .SYNOPSIS
            Provides a way to expand collections in an object property prior
            to being sent to Export-Csv.
        .DESCRIPTION
            Provides a way to expand collections in an object property prior
            to being sent to Export-Csv. This helps to avoid the object type
            from being shown such as system.object[] in a spreadsheet.
        .PARAMETER InputObject
            The object that will be sent to Export-Csv
        .PARAMETER OutPropertyType
            This determines whether the property that has the collection will be
            shown in the CSV as a comma delimmited string or as a stacked string.
            Possible values:
            Stack
            Comma
            Default value is: Stack
        .NOTES
            Name: Convert-OutputForCSV
            Author: Boe Prox
            Created: 24 Jan 2014
            Version History:
                1.1 - 02 Feb 2014
                    -Removed OutputOrder parameter as it is no longer needed; inputobject order is now respected 
                    in the output object
                1.0 - 24 Jan 2014
                    -Initial Creation
        .EXAMPLE
            $Output = 'PSComputername','IPAddress','DNSServerSearchOrder'
            Get-WMIObject -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" |
            Select-Object $Output | Convert-OutputForCSV | 
            Export-Csv -NoTypeInformation -Path NIC.csv    
            
            Description
            -----------
            Using a predefined set of properties to display ($Output), data is collected from the 
            Win32_NetworkAdapterConfiguration class and then passed to the Convert-OutputForCSV
            funtion which expands any property with a collection so it can be read properly prior
            to being sent to Export-Csv. Properties that had a collection will be viewed as a stack
            in the spreadsheet.        
            
    #>
    #Requires -Version 3.0
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline)]
        [psobject]$InputObject,
        [parameter()]
        [ValidateSet('Stack','SemiColon')]
        [string]$OutputPropertyType = 'Stack'
    )
    Begin {
        $PSBoundParameters.GetEnumerator() | ForEach {
            Write-Verbose "$($_)"
        }
        $FirstRun = $True
    }
    Process {
        If ($FirstRun) {
            $OutputOrder = $InputObject.psobject.properties.name
            Write-Verbose "Output Order:`n $($OutputOrder -join ', ' )"
            $FirstRun = $False
            #Get properties to process
            $Properties = Get-Member -InputObject $InputObject -MemberType *Property
            #Get properties that hold a collection
            $Properties_Collection = @(($Properties | Where-Object {
                $_.Definition -match "Collection|\[\]"
            }).Name)
            #Get properties that do not hold a collection
            $Properties_NoCollection = @(($Properties | Where-Object {
                $_.Definition -notmatch "Collection|\[\]"
            }).Name)
            Write-Verbose "Properties Found that have collections:`n $(($Properties_Collection) -join ', ')"
            Write-Verbose "Properties Found that have no collections:`n $(($Properties_NoCollection) -join ', ')"
        }
 
        $InputObject | ForEach {
            $Line = $_
            $stringBuilder = New-Object Text.StringBuilder
            $Null = $stringBuilder.AppendLine("[pscustomobject] @{")

            $OutputOrder | ForEach {
                If ($OutputPropertyType -eq 'Stack') {
                    $Null = $stringBuilder.AppendLine("`"$($_)`" = `"$(($line.$($_) | Out-String).Trim())`"")                 
                } 
                ElseIf ($OutputPropertyType -eq "SemiColon") {
                    $Null = $stringBuilder.AppendLine("`"$($_)`" = `"$($line.$($_) -join '; ')`"")                   
                }

            }
            $Null = $stringBuilder.AppendLine("}")
 
            Invoke-Expression $stringBuilder.ToString()
        }
    }
    End {}
}

#region // User-Specific Environmental Variables for NSX-T

    #Replace the following variables with your own environment
    $nsxserver = 'nsx01.projectonestone.io'
    $username = 'admin'
    $password = 'VMware1!VMware1!'

    #Do not modify the following four lines
    $auth = $username + ':' + $password

    $Encoded = [System.Text.Encoding]::UTF8.GetBytes($auth)
    $EncodedPassword = [System.Convert]::ToBase64String($Encoded)

    $headers = @{"Authorization"="Basic $($EncodedPassword)"}

#endregion

#region // Variables for file creation

    $dfwexport = "$env:userprofile\desktop\dfwexport.csv"
    $dfwexport1 = "$env:userprofile\desktop\dfwexport1.csv"
    $dfwexport2 = "$env:userprofile\desktop\dfwexport2.csv"
    $nsgroups = "$env:userprofile\desktop\nsgroups.csv"
    $nsgroupvmmembers = "$env:userprofile\desktop\nsgroupsvmmembers.csv"
    $nsgroupipmembers = "$env:userprofile\desktop\nsgroupsipmembers.csv"
    $nsservices = "$env:USERPROFILE\desktop\nsx-services.csv"
    $mergedxlsx = "$env:userprofile\desktop\DFW-Export.xlsx"

#endregion 

#region // Clean up stale files from previous deployment

    Remove-Item -Path $dfwexport -Confirm:$false -Force -ErrorAction Ignore
    Remove-Item -Path $dfwexport1 -Confirm:$false -Force -ErrorAction Ignore
    Remove-Item -Path $dfwexport2 -Confirm:$false -Force -ErrorAction Ignore
    Remove-Item -Path $nsgroups -Confirm:$false -Force -ErrorAction Ignore
    Remove-Item -Path $nsservices -Confirm:$false -Force -ErrorAction Ignore 
    Remove-Item -Path $mergedxlsx -Confirm:$false -Force -ErrorAction Ignore
    Remove-Item -Path $nsgroups -Confirm:$false -Force -ErrorAction Ignore
    Remove-Item -Path $nsgroupvmmembers -Confirm:$false -Force -ErrorAction Ignore
    Remove-Item -Path $nsgroupipmembers -Confirm:$false -Force -ErrorAction Ignore

#endregion

#region // Connect to the NSX-T Manager using PowerCLI

    Write-Host "`r`nConnecting to the NSX-T Manager Appliance.`r`nThis process takes approximately 1-2 min... " -ForegroundColor Yellow
    Connect-NsxtServer -server $nsxserver -User $username -Password $password
    Write-Host "`nConnected to NSX-T Manager Appliance. Starting Export..." -ForegroundColor Yellow


#endregion

#region // DFW Rules Export
    
    Write-Host "`r`nExporting: DFW Rules..." -ForegroundColor Yellow


    #region // Variables to invoke RestAPI call to pull security policies and ruleset identifiers

        $nsxpolicies = Invoke-RestMethod -Method Get -Uri "https://$nsxserver/policy/api/v1/infra/domains/default/security-policies" -Header $headers
        $nsxdfwrules = foreach ($id in $nsxpolicies.results) {Invoke-RestMethod -Method Get -Uri "https://$nsxserver/policy/api/v1/infra/domains/default/security-policies/$($id.id)/rules" -Header $headers}

    #endregion 


    #region // ForEach Loop to pull DFW rule sets based on Security Policy ID from previous command

        ForEach ($rule in $nsxdfwrules.results) {

            $output = 'id','display_name','description','source_groups','destination_groups','services','scope','action'
            $rule | Select-Object -Property $output | Convert-OutputForCSV -OutputPropertyType SemiColon | Export-Csv -Path $dfwexport -NoTypeInformation -Encoding UTF8 -Append
        } 

    #endregion

#endregion

#region // NSGroup Export
Write-Host "Exporting: NSX Security Groups and VM Members..." -ForegroundColor Yellow

    #region // Variables to invoke RestAPI call to pull NSX Security Groups (NSgroups)
        
        $nsxgroups = Invoke-RestMethod -Method Get -Uri "https://$nsxserver/policy/api/v1/infra/domains/default/groups" -Header $headers

    #endregion

    #region // ForEach Loop to pull NSgroups

        ForEach ($group in $nsxgroups.results) {

            $output = 'path','display_name'
            $group | Select-Object -Property $output | Convert-OutputForCSV -OutputPropertyType SemiColon | Export-Csv -Path $nsgroups -NoTypeInformation -Encoding UTF8 -Append

        }

    #endregion

#endregion

<##region // NSX Security Group Members Export for DFW rules

    #region // NSX Group - VM Members

            $domain_id = "default"     
            $ns_groupsvc = Get-NsxtPolicyService -Name com.vmware.nsx_policy.infra.domains.groups
            $ns_groups = $ns_groupsvc.list($domain_id)
            $ns_group = $ns_groups.results | ft -AutoSize -Property unique_id

            ForEach ($group_id in $ns_group.unique_id) {

                $vmsetsvc = Get-NsxtService -Name com.vmware.nsx.ns_groups.effective_virtual_machine_members 

                $vmsets = $vmsetsvc.list($group_id)
                            
                $vmsets.results | Select-Object -Property * | Convert-OutputForCSV -OutputPropertyType SemiColon | Export-Csv -path $nsgroupvmmembers -NoTypeInformation -Encoding UTF8 -Append

            }
        
    #endregion 

    #region // NSX Group - IP Members

        $ipsetsvc = Get-NsxtService -Name com.vmware.nsx.ip_sets
        $ipsets = $ipsetsvc.list()
        $ipsets.results | ft -AutoSize -Property id, ip_addresses
            
        ForEach ($ipset in $ipsets.results) {

                $output = 'id','ip_addresses'

                $ipset | Select-Object -Property $output | Convert-OutputForCSV -OutputPropertyType SemiColon | Export-Csv -Path $nsgroupipmembers -NoTypeInformation -Encoding UTF8 -Append
            
        }      

    #endregion 

#endregion#>

#region // NSX Services Export for DFW rules

    Write-Host "Exporting: NSX Service Definitions..." -ForegroundColor Yellow

    #region // Variables to invoke RestAPI call to pull NSX Security Groups (NSgroups)

        $nsxservices = Invoke-RestMethod -Method Get -Uri "https://$nsxserver/policy/api/v1/infra/services" -Header $headers 

    #endregion 

    #region // ForEach Loop to pull NSX Services

        ForEach ($service in $nsxservices.results) {

                ForEach ($service_entry in $service.service_entries) {
                        
                        $output = 'parent_path','display_name','path','id','l4_protocol','source_ports','destination_ports'
                        $service_entry | Select-Object -Property $output | Convert-OutputForCSV -OutputPropertyType SemiColon | Export-Csv -Path $nsservices -NoTypeInformation -Encoding UTF8 -Append
                } 
        }

    #endregion

#endregion

#region // Create Hash Table from NSGroups

    
    $import1 = Import-Csv $nsgroups -Header path,display_name -Encoding UTF8 -Delimiter ","
    
    $grouphash = @{}

    ForEach ($i in $import1) {

        $grouphash[$i.path]=$i.display_name       

    }
    
#endregion

#region // Find and Replace in DFW CSV from NSGroups Hash
    
    $file = Get-Content $dfwexport

    $file | ForEach-Object {

        $line = $_

        $grouphash.GetEnumerator() | ForEach-Object {

            if ($line -match $_.Key) {

                $line = $line -replace $_.Key, $_.Value

            }

        }

        $line

    } | Set-Content -Path $env:userprofile\desktop\dfwexport1.csv

#endregion 

#region // Create Hash Table from Services

    $import2 = Import-Csv $nsservices -Header path,display_name -Encoding UTF8 -Delimiter ","
    
    $servicehash = @{}

    ForEach ($i in $import2) {

        $servicehash[$i.path]=$i.display_name       

    }
    
#endregion

#region // Find and Replace in DFW CSV from Services Hash
    
    $file2 = Get-Content $env:userprofile\desktop\dfwexport1.csv

    $file2 | ForEach-Object {

        $line = $_

        $servicehash.GetEnumerator() | ForEach-Object {

            if ($line -match $_.Key) {

                $line = $line -replace $_.Key, $_.Value

            }

        }

        $line

    } | Set-Content -Path $env:userprofile\desktop\dfwexport2.csv

#endregion 

#region // Remove Old DFW Export
    
    Remove-Item -Path $dfwexport -Confirm:$false -Force -ErrorAction Ignore
    Remove-Item -Path $dfwexport1 -Confirm:$false -Force -ErrorAction Ignore
    Rename-Item -Path $dfwexport2 -NewName $dfwexport -ErrorAction Ignore
    Write-Host "`r`nExport Complete..." -ForegroundColor Green

#endregion

#region // Merge CSV Files into a single Excel Workbook

    $path = "$env:userprofile\desktop"

    $csvs = Get-ChildItem $path\* -Include *.csv
    $i=$csvs.Count
        
    Write-Host "`r`nDetected the following CSV files: ($i)" -ForegroundColor Yellow

        foreach ($csv in $csvs) {
            Write-Host " "$csv.Name
        }

    $outputfilename = "DFW-Export.xlsx"

    Write-Host "`r`nCreating: $outputfilename..." -ForegroundColor Yellow

    $excelapp = new-object -comobject Excel.Application
    $excelapp.sheetsInNewWorkbook = $csvs.Count
    $xlsx = $excelapp.Workbooks.Add()
    $sheet=1

    foreach ($csv in $csvs)
    {
        $row=1
        $column=1
        $worksheet = $xlsx.Worksheets.Item($sheet)
        $worksheet.Name = $csv.Name
        $file = (Get-Content $csv)
        foreach ($line in $file) {
            $linecontents=$line -split ',(?!\s*\w+")'

            foreach($cell in $linecontents) {

                $worksheet.Cells.Item($row,$column) = $cell
                $column++

            }

            $column=1
            $row++

        }
            
        $sheet++

    }

    $output = "$($path)\$($outputfilename)"
    $xlsx.SaveAs($output)
    $excelapp.quit()
    Write-Host "`r`nObject: $($outputfilename) created in $($path) on $(get-date -f yyyy-MM-ddTHH:mm:ss:ff)" -ForegroundColor Green

#endregion

#region // Clean up files and Disconnect from NSX-T

    Write-Host "`nCleaning up files..." -ForegroundColor White
    
    Remove-Item -Path $dfwexport -Confirm:$false -Force -ErrorAction Ignore
    Remove-Item -Path $nsgroups -Confirm:$false -Force -ErrorAction Ignore
    Remove-Item -Path $nsservices -Confirm:$false -Force -ErrorAction Ignore 
    Remove-Item -Path $nsgroupvmmembers -Confirm:$false -Force -ErrorAction Ignore
    Remove-Item -Path $nsgroupipmembers -Confirm:$false -Force -ErrorAction Ignore
    
    Write-Host "`nDisconnecting NSX-T Servers" -ForegroundColor Yellow

    Disconnect-NsxtServer -Server * -Confirm:$false

    Write-Host "`nProcess Complete." -ForegroundColor Green

#endregion
