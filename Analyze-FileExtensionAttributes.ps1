<#
This code is functional, but
it needs quite a bit of polish
for usability.

#TO-DOs:
    * Add input parameters
    * Control/define output file name
#>

#region Prerequisites - Variables, import files

$Folder = "D:\Extended Attrs Compilation Files\"
$ShellObj = New-Object -ComObject Shell.Application
$Data = New-Object System.Collections.ArrayList

$ReclaimMemory = {
    
    [gc]::Collect()
    [System.GC]::GetTotalMemory('forcefullcollection') | Out-Null
    [System.GC]::GetTotalMemory($true) | out-null

} #Close Scriptblock ReclaimMemory

$CSVFiles = get-childitem $Folder *.csv
$CSVCount = $CSVFiles.Count

$w = 1

$CSVFiles.ForEach({
    
    $CSV = $_
    
    Write-Progress -Activity "Importing CSVs" -Status "Getting $($CSV.Name) | $($CSVCount) entries" -Id 1 -PercentComplete ([int](($w / $CSVCount) * 100))    
    
    (Import-csv ($CSV.FullName)).ForEach({$Data.Add($_) | out-null})
    
    $w++

})

$Extensions = $Data | Group-Object 'File extension'

Remove-Variable Data
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

$Results = New-Object System.Collections.ArrayList

$ExtCount = $Extensions.Count
$x = 1

#endregion

Foreach ($Extension in $Extensions){

$ExtData = $Extension.Group
$PropertiesHash = @{}
$TotalProps = @{} 

$FileCount = $ExtData.Count
$y = 1

Write-Progress -Activity "Analyzing Extensions" -Status "Working on $($Extension.Name) | $($ExtData.Count) entries" -Id 1 -PercentComplete ([int](($x / $ExtCount) * 100))

#region determine unique properties
Foreach ($Line in $ExtData){

    Write-Progress -Activity "Scanning for unique property names" -Status "$($Line.Name)" -ParentId 1 -PercentComplete ([int](($y / $FileCount) * 100))

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

$ExtData.ForEach({
    
    Write-Progress -Activity "Checking property values" -Status "$($_.Name))" -ParentId 1 -PercentComplete ([int](($y / $FileCount) * 100))
    
    $ObjProperties = $_.psobject.Properties.Name
  
    :IndexLoop Foreach ($Index in (0..$ObjProperties.GetUpperBound(0))){
  
        $Property = $ObjProperties[$Index]
        
        Try{[System.Collections.arraylist]$PropertyHashValues = $PropertiesHash."$Property"}
        Catch {} #Suppress "Length" named property error
        
          If ($null -eq $PropertyHashValues){$PropertiesHash.Add($Property,[system.collections.arraylist]@($Index))}
          
          Else{
            
            $PropertyHashValues.Add($Index) | Out-Null
            
            $SortedValues = New-Object System.Collections.ArrayList
            ($PropertyHashValues | Sort-Object -Unique -Descending).ForEach({$SortedValues.Add($_) | out-null})
            
            $PropertiesHash.Remove($Property) | out-null
  
            $PropertiesHash.Add($Property,$SortedValues)
  
          }
  
          If ($PropertiesHash.Count -eq $TotalPropsCount){break IndexLoop} #As soon as we've encountered every possible property, break
  
        } 

        $y++

    })
  
    $PropertyArrangement = New-Object System.Collections.ArrayList
  
    ($PropertiesHash.Keys).foreach({$PropertyArrangement.Add([pscustomobject]@{"Name" = $_; "Index" = $(($PropertiesHash.$_)[0])}) | Out-Null })
  
    $AllProperties = ($PropertyArrangement | Sort-Object "Index").Name
  
    $UsedProperties = New-Object System.Collections.ArrayList
  
    Remove-Variable PropertiesHash,TotalPropsCount,PropertyArrangement -ErrorAction SilentlyContinue
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

Remove-Variable PropertiesHash,TotalPropsCount,PropertyArrangement -ErrorAction SilentlyContinue

  &$ReclaimMemory

$x++

#endregion

}

$Results | ConvertTo-Json | out-file "exthelper.json"