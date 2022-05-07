Function Get-FileExtension([string]$Path){$P = $Path.Split('.'); return "$('.' + $P[$P.Count -1])"}

Function Get-Files ($Directory,[switch]$ExcludeFullPath){

Try {$Files = [System.IO.Directory]::EnumerateFiles("$Directory","*.*","TopDirectoryOnly")}
Catch {$_.Exception; Continue}

$FilesNormalized = New-Object System.Collections.ArrayList
$Files.Where({$_ -notmatch 'db'}).ForEach({$FilesNormalized.Add($_) | Out-Null})

If ($ExcludeFullPath.IsPresent -and $FilesNormalized.Count -ne 0){ 
    
    (0..($FilesNormalized.Count - 1 )).ForEach({
        
        $FilesNormalized[$_] = $FilesNormalized[$_].Replace("$Directory",'').TrimStart('\\')
    
    }) 

}

Return $FilesNormalized

} #Close Function

#Function Get-StringHash ($String){Return (Get-FileHash -InputStream ([IO.MemoryStream]::new([byte[]][char[]]$String)) -Algorithm SHA1).Hash}

Function Get-Directories ($Directory){
    $Dirs = New-Object System.Collections.ArrayList

    Try {([System.IO.Directory]::EnumerateDirectories("$Directory","*","TopDirectory")).ForEach({$Dirs.Add($_) | Out-Null})}
    catch {throw "Unable to enumerate directories in $Directory"}

    Return $Dirs

}

Function Get-SubDirectories ($Directory,[switch]$SuppressErrors){

$SubDirs = New-Object System.Collections.ArrayList

If ($Directory.Length -eq 0){$Directory = (Get-Location).ProviderPath}

[System.Collections.ArrayList]$Directories = (Get-Directories -Directory $Directory).Where({$_ -inotmatch 'filehistory|windows|recycle|@'})

$DirCount = $Directories.Count

Do {

    $DirQueue = New-Object System.Collections.ArrayList
    
    :DirLoop Foreach ($Dir in $Directories){
        
        $SubDirs.Add($Dir) | Out-Null
        
        Try {$DirQueue += (Get-Directories -Directory $Dir)}
        Catch {
            
            If (!$SuppressErrors.IsPresent){Write-Error "Cannot enumerate directories in $D"}
            Continue DirLoop
        
        }
        
    }

    $Directories = $DirQueue

    $DirCount = $Directories.Count

} #Close Do

Until ($DirCount -eq 0)

return $SubDirs

} #Close Get SubDirectories

Function Get-ExtendedAttributes {

[CmdletBinding()] 
param( 
    [Parameter(Mandatory=$False)] [string]$Path=((Get-Location).ProviderPath), 
    [Parameter(Mandatory=$False)][switch]$Recurse,
    [Parameter(Mandatory=$False)][switch]$WriteProgress,
    [Parameter(ParameterSetName='HelperFile',Mandatory=$False)] [switch]$UseHelperFile,
    [Parameter(ParameterSetName='HelperFile',Mandatory=$False)] [string]$HelperFilename="exthelper.json",
    [Parameter(Mandatory=$False)][array]$FilterOut,
    [Parameter(Mandatory=$False)][array]$IncludeOnly,
    [Parameter(Mandatory=$False)][switch]$ReduceDown
)

begin {
#region Define preemptive variables
$ShellObj = New-Object -ComObject Shell.Application
$Results = New-Object System.Collections.ArrayList

$ReclaimMemory = {
    
    [gc]::Collect()
    [System.GC]::GetTotalMemory('forcefullcollection') | Out-Null
    [System.GC]::GetTotalMemory($true) | out-null

} #Close Scriptblock ReclaimMemory

#endregion

#region Build Filter expression
If ($FilterOut.Count -gt 0){$OutFilter = $FilterOut -join '|'; $OutFilterEnabled = $true}
Else {$OutFilterEnabled = $false}

If ($IncludeOnly.Count -gt 0){$InFilter = $IncludeOnly -join '|'; $InFilterEnabled = $true}
Else {$InFilterEnabled = $false}

#endregion

#region Enumerate Attribute Columns
Try {$RootPath = $ShellObj.Namespace($Path)}
Catch {throw "Unable to initialize Shell Application for namespace $Path"}

$AttrHash = @{}

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

$ValidColumns = New-Object System.Collections.ArrayList

($AttrHash.Keys | Sort-Object).ForEach({$ValidColumns.Add($_) | Out-Null})
$NoValueAttrs.ForEach({$ValidColumns.Remove($_)})

#endregion 

#region Import Helper File into Helper Hash Table:
If ($UseHelperFile.IsPresent){
    $HelperHash = @{}

    Try {$JSON = Get-Content $HelperFilename -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop}
    Catch {throw "Helper file $HelperFilename is not valid"}

    $JSON.ForEach({
                    
                    [System.Collections.ArrayList]$HelperAttrs = $_.Attrs
                    $NoValueAttrs.ForEach({$HelperAttrs.Remove($_)}) #Remove "No-Value Attributes"
                    $HelperHash.Add(($_.Extension),($_.Attrs))
    
                })

    Remove-Variable JSON -ErrorAction SilentlyContinue
    &$ReclaimMemory

} #Close If UseHelper.IsPresent

#endregion

} #Close Begin

Process {

#region Build Directory/File Hashtable

$DirIndex = @{}

$DirIndex.Add("$Path",(Get-Files -Directory $Path -ExcludeFullPath))

If ($Recurse.IsPresent){
    
      $SubDirs = (Get-SubDirectories -Directory $Path -SuppressErrors)
      
      Foreach ($Dir in $SubDirs){$DirIndex.Add("$Dir",(Get-Files -Directory $Dir -ExcludeFullPath))}
      
      Remove-Variable SubDirs -ErrorAction SilentlyContinue
      &$ReclaimMemory

} #Close If Recurse.IsPresent

$KeyDirs = $DirIndex.GetEnumerator().Name | Out-String -Stream

If ($OutFilterEnabled -eq $true){$KeyDirs = $KeyDirs.Where({$_ -inotmatch "$OutFilter"})}

    #Commented this out because "InFilter" works best against full file paths:
#If ($InFilterEnabled -eq $true) {$KeyDirs = $KeyDirs.Where({$_ -imatch "$InFilter"})}

#endregion

#region Filter hash table file values
$KeyDirs.ForEach({
    $Dir = $_

    If ($OutFilterEnabled -eq $true){$DirIndex."$Dir" = $DirIndex."$Dir".Where({$_ -inotmatch "$OutFilter"})}
    If ($InFilterEnabled -eq $true) {$DirIndex."$Dir" = $DirIndex."$Dir".Where({$_ -imatch "$InFilter"})}

}) 

#endregion

#region Gather file metadata

$FileCount = 0
$KeyDirs.ForEach({$FileCount += ($DirIndex.$_).Count})

$x = 1 #Counter

:KeyLoop Foreach ($Dir in $KeyDirs){

$Files = $DirIndex."$Dir"

If ($Files.Count -eq 0){Continue KeyLoop}

$Files.ForEach({

    $File = $_

    $FolderObj = $ShellObj.NameSpace($Dir)
    $FileObj = $FolderObj.ParseName($File)

    If ($WriteProgress.IsPresent){Write-Progress -Activity "Retrieving Extended Attributes" -Status "Working on $Dir\$File" -PercentComplete ([int](($x / $FileCount) * 100))}

    $Object = New-Object System.Object

    Switch ($UseHelperFile.IsPresent){

    $True {

          $FileAttrs = $HelperHash."$(Get-FileExtension -Path $File)"

          If ($FileAttrs.Count -gt 0){$TargetColumns = $FileAttrs}
          Else {$TargetColumns = $ValidColumns}

    }#Close True

    $False {$TargetColumns = $ValidColumns}

    }#Close Switch UserHelperFile.IsPresent

    $TargetColumns.ForEach({ 
        
        Try {$Object | Add-Member -ErrorAction Stop -Membertype Noteproperty -Name "$($AttrHash.$_)" -Value ($FolderObj.GetDetailsOf($FileObj, $_))}
        Catch {$Object | Add-Member -ErrorAction Stop -Membertype Noteproperty -Name "$($AttrHash.$_)" -Value "-"; Continue}
    
        })

    $Results.Add($Object) | Out-Null

    $x++ 

    }) #Close Files.ForEach

} #Close :KeyLoop

#endregion

} #Close Process

end {
    
    #region Filter out unused properties
If ($ReduceDown.IsPresent){
        $PropertiesHash = @{}
        $Properties = $Results[0].psobject.Properties.Name
        $Properties.ForEach({$PropertiesHash.Add($_,0)})
        $UsedProperties = New-Object System.Collections.ArrayList

        $Results.ForEach({
        $Obj = $_
        $Properties.ForEach({If ($Obj.$_.Length -ge 1){$PropertiesHash.$_++}})
        })

        $Properties.ForEach({If ($PropertiesHash.$_ -ge 1){$UsedProperties.Add($_) | Out-Null}})

        $Results = $Results | Select-Object $UsedProperties

} 

Else {$Results = $Results | Select-Object *}

Return $Results

    #endregion

} #Close End

} #Close Function