## Convert-OutputForCSV is created by Boe Prox (GITHUB proxb - https://github.com/proxb/PowerShell_Scripts/blob/master/Convert-OutputForCSV.ps1); all credit goes to him for the Convert-OutputForCSV function)

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
        [ValidateSet('Stack','Comma')]
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
                } ElseIf ($OutputPropertyType -eq "Comma") {
                    $Null = $stringBuilder.AppendLine("`"$($_)`" = `"$($line.$($_) -join ', ')`"")                   
                }
            }
            $Null = $stringBuilder.AppendLine("}")
 
            Invoke-Expression $stringBuilder.ToString()
        }
    }
    End {}
}

#Replace the following variables with your own environment
$nsxserver = 'https://nsx01.projectonestone.io'
$username = 'admin'
$password = 'VMware1!VMware1!'


#Begin Script

$dfwexport = "$env:userprofile\desktop\dfwexport.csv"
$dfwexport1 = "$env:userprofile\desktop\dfwexport1.csv"
$dfwexport2 = "$env:userprofile\desktop\dfwexport2.csv"
$nsgroups = "$env:userprofile\desktop\nsgroups.csv"
$nsgroupvmmembers = "$env:userprofile\desktop\nsgroupsvmmembers.csv"
$nsgroupipmembers = "$env:userprofile\desktop\nsgroupipmembers.csv"
$nsservices = "$env:USERPROFILE\desktop\nsx-services.csv"
$mergedxlsx = "$env:userprofile\desktop\DFW-Export.xlsx"


Remove-Item -Path $dfwexport -Confirm:$false -Force -ErrorAction Ignore
Remove-Item -Path $dfwexport1 -Confirm:$false -Force -ErrorAction Ignore
Remove-Item -Path $dfwexport2 -Confirm:$false -Force -ErrorAction Ignore
Remove-Item -Path $nsgroups -Confirm:$false -Force -ErrorAction Ignore
Remove-Item -Path $nsservices -Confirm:$false -Force -ErrorAction Ignore 
Remove-Item -Path $mergedxlsx -Confirm:$false -Force -ErrorAction Ignore
Remove-Item -Path $nsgroups -Confirm:$false -Force -ErrorAction Ignore
Remove-Item -Path $nsgroupvmmembers -Confirm:$false -Force -ErrorAction Ignore
Remove-Item -Path $nsgroupipmembers -Confirm:$false -Force -ErrorAction Ignore

$auth = $username + ':' + $password

$Encoded = [System.Text.Encoding]::UTF8.GetBytes($auth)
$EncodedPassword = [System.Convert]::ToBase64String($Encoded)

$headers = @{"Authorization"="Basic $($EncodedPassword)"}



 
 #region // DFW Rules Export
 Write-Host "Exporting: DFW Rules..." -ForegroundColor Yellow


    #region // Variables to invoke RestAPI call to pull security policies and ruleset identifiers

        $nsxpolicies = Invoke-RestMethod -Method Get -Uri "$nsxserver/policy/api/v1/infra/domains/default/security-policies" -Header $headers
        $nsxdfwrules = foreach ($id in $nsxpolicies.results) {Invoke-RestMethod -Method Get -Uri "$nsxserver/policy/api/v1/infra/domains/default/security-policies/$($id.id)/rules" -Header $headers}

    #endregion 


    #region // ForEach Loop to pull DFW rule sets based on Security Policy ID from previous command

        ForEach ($rule in $nsxdfwrules.results) {

            $output = 'id','display_name','description','source_groups','destination_groups','services','scope','action'
            $rule | Select-Object -Property $output | Convert-OutputForCSV -OutputPropertyType Comma | Export-Csv -Path $dfwexport -NoTypeInformation -Encoding UTF8 -Append
        } 

    #endregion

#endregion



#region // NSGroup Export
Write-Host "Exporting: NSX Security Groups and VM Members..." -ForegroundColor Yellow

    #region // Variables to invoke RestAPI call to pull NSX Security Groups (NSgroups)
        
        $nsxgroups = Invoke-RestMethod -Method Get -Uri "$nsxserver/policy/api/v1/infra/domains/default/groups" -Header $headers

    #endregion

    #region // ForEach Loop to pull NSgroups

        ForEach ($group in $nsxgroups.results) {

            $output = 'path','display_name'
            $group | Select-Object -Property $output | Convert-OutputForCSV -OutputPropertyType Comma | Export-Csv -Path $nsgroups -NoTypeInformation -Encoding UTF8 -Append

        }

    #endregion

#endregion



#region // NSX Security Group Members Export for DFW rules

    #region // NSX Group - VM Members

        ForEach ($group in $nsxgroups.results) {

            $groupid = $group
                               
            $nsxgroupmembers = Invoke-RestMethod -Method Get -Uri "$nsxserver/policy/api/v1/infra/domains/default/groups/$($groupid.id)/members/virtual-machines" -Header $headers 

            ForEach ($groupmember in $nsxgroupmembers.results) {

                    $output = 'path','host_id','local_id','display_name'     
                    $groupmember | Select-Object -Property $output | Convert-OutputForCSV -OutputPropertyType Comma | Export-Csv -Path $nsgroupvmmembers -NoTypeInformation -Encoding UTF8 -Append
            
            }
        }
        
    #endregion 

    <# STILL IN DEVELOPMENT 

    #region // NSX Group - IP Members

        ForEach ($group in $nsxgroups.results) {

            $groupid = $group

            $nsxgroupmembers = Invoke-RestMethod -Method Get -Uri "$nsxserver/policy/api/v1$($groupid.path)/members/ip-addresses" -Header $headers 

            ForEach ($ipgroupmember in $nsxgroupmembers.results) {

                   $ipgroupmember | Convert-OutputForCSV -OutputPropertyType Comma | Export-Csv -Path $nsgroupipmembers -NoTypeInformation -Append
            
            }          
            
        }

    #endregion 
    #>

#endregion







#region // NSX Services Export for DFW rules
Write-Host "Exporting: NSX Service Definitions..." -ForegroundColor Yellow

    #region // Variables to invoke RestAPI call to pull NSX Security Groups (NSgroups)

        $nsxservices = Invoke-RestMethod -Method Get -Uri "$nsxserver/policy/api/v1/infra/services" -Header $headers 

    #endregion 

    #region // ForEach Loop to pull NSX Services

        ForEach ($service in $nsxservices.results) {

                ForEach ($service_entry in $service.service_entries) {
                        
                        $output = 'parent_path','display_name','path','id','l4_protocol','source_ports','destination_ports'
                        $service_entry | Select-Object -Property $output | Convert-OutputForCSV -OutputPropertyType Comma | Export-Csv -Path $nsservices -NoTypeInformation -Encoding UTF8 -Append
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
    #Remove-Item -Path $nsgroups -Confirm:$false -Force -ErrorAction Ignore
    Rename-Item -Path $dfwexport2 -NewName $dfwexport -ErrorAction Ignore

#endregion


#region // Merge CSV Files into a single Excel Workbook

    $path = "$env:userprofile\desktop"

    $csvs = Get-ChildItem $path\* -Include *.csv
    $i=$csvs.Count
        
    Write-Host "Detected the following CSV files: ($i)" -ForegroundColor Green

    foreach ($csv in $csvs) {
        Write-Host " "$csv.Name
    }

    $outputfilename = "DFW-Export.xlsx"

    Write-Host "Creating: $outputfilename" -ForegroundColor Yellow

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
    Write-Host "Object: $($outputfilename) created in $($path) on $(get-date -f yyyy-MM-ddTHH:mm:ss:ff)" -ForegroundColor Yellow

#endregion

    Write-Host "Cleaning up files..." -ForegroundColor White

    Remove-Item -Path $dfwexport -Confirm:$false -Force -ErrorAction Ignore
    Remove-Item -Path $nsgroups -Confirm:$false -Force -ErrorAction Ignore
    Remove-Item -Path $nsservices -Confirm:$false -Force -ErrorAction Ignore 
    Remove-Item -Path $nsgroupvmmembers -Confirm:$false -Force -ErrorAction Ignore
    Remove-Item -Path $nsgroupipmembers -Confirm:$false -Force -ErrorAction Ignore

    Write-Host "Process Complete." -ForegroundColor Yellow

#End Script
