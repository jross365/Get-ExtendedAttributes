
    #WorkBench.ps1 is for cleaning up/testing changes to functions and blocks of code



<#
This code is functional, but
it needs quite a bit of polish
for usability.

#TO-DOs:
    * Add input parameters
    * Control/define output file name
#>

#region Prerequisites - Variables, import files

function New-AttrsHelperFile {
[CmdletBinding()] 
    param( 
        [Parameter(Mandatory=$True)] [string]$Folder, 
        [Parameter(Mandatory=$True)] [string]$SaveAs,
        [Parameter(Mandatory=$False)] [switch]$WriteProgress
        
    )

begin {
#region Define and check variables
If (!(Test-Path $Folder)){throw "$Folder is not a valid path"}

$Results = New-Object System.Collections.ArrayList
$ShellObj = New-Object -ComObject Shell.Application
$Data = New-Object System.Collections.ArrayList

$ReclaimMemory = {
    
    [gc]::Collect()
    [System.GC]::GetTotalMemory('forcefullcollection') | Out-Null
    [System.GC]::GetTotalMemory($true) | out-null

} #Close Scriptblock ReclaimMemory

$Beep = {[system.console]::Beep()}

$WrProg = {$WriteProgress.IsPresent -eq $true}

$CSVFiles = Get-Files -Directory $Folder -Filter .csv
$CSVCount = $CSVFiles.Count

If ($CSVCount -eq 0){throw "No CSVs were found in $Folder"}

If (Test-Path $SaveAs){throw "$SaveAs already exists, please specify a new and unique file name"}

#define counters for top-level loops:
$w = 1
$x = 1

#endregion

#region import CSVs, group results, clean up memory

$CSVFiles.ForEach({
    
    $CSV = $_
    
    If (&$WrProg){Write-Progress -Activity "Importing CSVs" -Status "Getting $CSV | $CSVCount entries" -Id 1 -PercentComplete ([int](($w / $CSVCount) * 100))}
    
    (Import-csv ($CSV)).ForEach({$Data.Add($_) | out-null})
    
    $w++

})

If (&$WrProg){Write-Progress -Activity "Importing CSVs" -Status "Completed" -Id 1 -Completed}

$Extensions = $Data | Group-Object 'File extension'

&$ReclaimMemory

#endregion

#region Enumerate Attribute Columns
$RootPath = $ShellObj.Namespace("$Folder")

$AttrHash = @{}
$ReverseAttrHash = @{}

$ValidColumns = New-Object System.Collections.ArrayList

(0..499).foreach({

$ColName = ($RootPath.GetDetailsOf($RootPath.Items, $_))

If ($ColName.Length -gt 0){$AttrHash.Add($_,$ColName)}

})

#These are duplicate metadata values:
$AttrHash.36  = 'Masters keywords[1]'
$AttrHash.149 = 'Status[1]'
$AttrHash.304 = 'Status[2]'

#These are low-value/duplicate attributes:
$NoValueAttrs = @(2,29,57,61,169,192,196,202,254,295)

($AttrHash.Keys | Sort-Object).ForEach({$ValidColumns.Add($_) | Out-Null})

$ValidColumns.ForEach({$ReverseAttrHash.Add(($AttrHash.$_),($_))})

$ExtCount = $Extensions.Count

#endregion

}

process {

    Foreach ($Extension in $Extensions){

        #region grab or define prerequisite inputs/outputs
        $ExtData = $Extension.Group
        $PropertiesHash = @{}
        $TotalProps = @{} 
        
        $FileCount = $ExtData.Count

        $y = 1
        
        If (&$WrProg){Write-Progress -Activity "Analyzing Extensions" -Status "Working on $($Extension.Name) | $($ExtData.Count) entries" -Id 1 -PercentComplete ([int](($x / $ExtCount) * 100))}
        
        #endregion

        #region determine unique properties
        Foreach ($Line in $ExtData){
        
            If (&$WrProg){Write-Progress -Activity "Scanning for unique property names" -Status "$($Line.Name)" -ParentId 1 -PercentComplete ([int](($y / $FileCount) * 100))}
        
            ($Line.psobject.Properties.Name).ForEach({
                
                $Property = $_
            
                If ($null -eq ($TotalProps."$Property")){$TotalProps.Add("$Property",0)}
          
                })
        
                $y++
           
            }
            
            $TotalPropsCount = $TotalProps.Count
            
            Remove-Variable TotalProps; &$ReclaimMemory
        
        #endregion    
        
        #region Inventory all properties in use by files with this extension
        
        $y = 1
        
    :PropLoop Foreach ($Line in $ExtData){
            
            If (&$WrProg){Write-Progress -Activity "Checking property values" -Status "$($Line.Name)" -ParentId 1 -PercentComplete ([int](($y / $FileCount) * 100))}
            
            $ObjProperties = $Line.psobject.Properties.Name
            
            $Index = 0
            $ObjPropsCount = $ObjProperties.GetUpperBound(0)

            Do {

                $Property = $ObjProperties[$Index]
                
                Try{[System.Collections.arraylist]$PropertyHashValues = $PropertiesHash."$Property"}
                Catch {} #Suppress "Length" named property error
                
                  If ($null -eq $PropertyHashValues){$PropertiesHash.Add($Property,[system.collections.arraylist]@($Index))}
                  Else {
                    
                    $PropertyHashValues.Add($Index) | Out-Null
                    
                    $SortedValues = New-Object System.Collections.ArrayList
                    ($PropertyHashValues | Sort-Object -Unique -Descending).ForEach({$SortedValues.Add($_) | out-null})
                    
                    $PropertiesHash.Remove($Property) | out-null
          
                    $PropertiesHash.Add($Property,$SortedValues)
          
                  }

                  $Index++


            }
            Until (($PropertiesHash.Count -eq $TotalPropsCount) -or ($Index -eq $ObjPropsCount))

            #If we have all of the properties available, Break the "Master" loop:
            If ($PropertiesHash.Count -eq $TotalPropsCount){Break PropLoop} 

                $y++
        
            }
          
            $PropertyArrangement = New-Object System.Collections.ArrayList
          
            ($PropertiesHash.Keys).foreach({$PropertyArrangement.Add([pscustomobject]@{"Name" = $_; "Index" = $(($PropertiesHash.$_)[0])}) | Out-Null })
          
            $AllProperties = ($PropertyArrangement | Sort-Object "Index").Name
          
            $UsedProperties = New-Object System.Collections.ArrayList

            &$ReclaimMemory
          
            $AllProperties.ForEach({
          
            $Property = $_
          
            If (($ExtData.Where({$_."$Property".Length -ne 0})).Count -gt 0){$UsedProperties.Add($Property) | Out-Null}
                
            })
        
        #endregion
        
        #region Look up Attribute numbers from names, filter out "NoValueAttrs"
        
        $UsedPropertyNos = New-Object System.Collections.ArrayList
        $UsedProperties.ForEach({$UsedPropertyNos.Add($ReverseAttrHash."$_") | Out-Null})
        $NoValueAttrs.ForEach({$Attr = $_; $UsedPropertyNos = $UsedPropertyNos.Where({$_ -ne $Attr})})
        
        
        $Object = New-Object System.Object
        $Object | Add-Member -MemberType Noteproperty -Name "Extension" -Value "$($Extension.Name)"
        $Object | Add-Member -MemberType NoteProperty -Name "Attrs" -Value ($UsedPropertyNos)
        $Results.Add($Object) | out-null
        
        
        $Extensions = $Extensions.Where({$_.Name -ne ($Extension.Name)})

        &$ReclaimMemory
        
        $x++
        
        #endregion
        
        } #Close Foreach $Extension

}

end {
   
    #region Final cleanup
    if (&$WrProg){
        Write-Progress -Activity "Checking property values" -Status "Complete" -ParentId 1 -Completed
        Write-Progress -Activity "Analyzing Extensions" -Status "Complete" -Id 1 -Completed
        }

    #endregion

    $Results | ConvertTo-Json | out-file "$SaveAs" -ErrorAction stop; 
    Write-Host "Results saved as $SaveAs"

    }

}











