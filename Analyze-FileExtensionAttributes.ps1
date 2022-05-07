$Folder = "D:\Extended Attrs Compilation Files\"
$ShellObj = New-Object -ComObject Shell.Application

(get-childitem $Folder *.csv).ForEach({$Data += import-csv ($_.FullName)})

$Extensions = $Data | Group-Object 'File extension'

Remove-Variable Data
[gc]::Collect()

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

($AttrHash.Keys | Sort-Object).ForEach({$ValidColumns.Add($_) | Out-Null})

$ValidColumns.ForEach({$ReverseAttrHash.Add(($AttrHash.$_),($_))})

$Results = New-Object System.Collections.ArrayList

$ExtCount = $Extensions.Count
$x = 1


Foreach ($Extension in $Extensions){


$ExtData = $Extension.Group

Write-Progress -Activity "Analyzing Extensions" -Status "Working on $($Extension.Name) | $($ExtData.Count) entries" -PercentComplete ([int](($x / $ExtCount) * 100))

$PropertiesHash = @{}
$Properties = $ExtData[0].psobject.Properties.Name
$Properties.ForEach({$PropertiesHash.Add($_,0)})
$UsedProperties = New-Object System.Collections.ArrayList

$ExtData.ForEach({
$Obj = $_
$Properties.ForEach({If ($Obj.$_.Length -ge 1){$PropertiesHash.$_++}})
})

$Properties.ForEach({If ($PropertiesHash.$_ -ge 1){$UsedProperties.Add($_) | Out-Null}})

$UsedPropertyNos = New-Object System.Collections.ArrayList
$UsedProperties.ForEach({$UsedPropertyNos.Add($ReverseAttrHash."$_") | Out-Null})

$Object = New-Object System.Object
$Object | Add-Member -MemberType Noteproperty -Name "Extension" -Value "$($Extension.Name)"
$Object | Add-Member -MemberType NoteProperty -Name "Attrs" -Value ($UsedPropertyNos)
$Results.Add($Object)

$Extensions = $Extensions.Where({$_.Name -ne ($Extension.Name)})

Remove-Variable ExtData
Remove-Variable PropertiesHash
Remove-Variable UsedPropertyNos
Remove-Variable Properties 
[gc]::Collect()
[System.GC]::GetTotalMemory('forcefullcollection') | Out-Null
[System.GC]::GetTotalMemory($true) | out-null

$x++

}

$Results | ConvertTo-Json | out-file "exthelper.json"