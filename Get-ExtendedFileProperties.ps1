<#
.SYNOPSIS
Enumerates directories (folders).

.DESCRIPTION
Enumerates and returns all directories in a provided path.
Can enumerate directories on at the top (root) level, or recursively.

.PARAMETER Directory
The top-level (root) directory to operate against.

.PARAMETER SuppressErrors
If specified, encountered errors (such as 'Access Denied') will not be reported.

.PARAMETER SuppressErrors
If specified, encountered errors (such as 'Access Denied') will not be reported.

.PARAMETER Recurse
If specified, will recursively enumerate all subdirectories nested in the root directory.

.PARAMETER NoSort
If specified, enumerated directories will be returned out-of-order. This improves the function's performance.

.PARAMETER IncludeRoot
If specified, the directories returned will include the top-level (root) directory.

.INPUTS
None. You cannot pipe objects to Get-Folders.

.OUTPUTS
System.Collections.ArrayList. Get-Folders returns an arraylist of string objects.

.EXAMPLE
Get-Folders C:\Users\User1\Desktop\Docs

C:\Users\User1\Desktop\Docs\Financial
C:\Users\User1\Desktop\Docs\Docs\Job Seeking
C:\Users\User1\Desktop\Docs\Docs\OBS Profiles
C:\Users\User1\Desktop\Docs\Docs\Personal
C:\Users\User1\Desktop\Docs\Docs\Recipes
C:\Users\User1\Desktop\Docs\Docs\Vehicle-Related

.EXAMPLE
Get-Folders C:\Users\User1\Desktop\Docs -Recurse

C:\Users\User1\Desktop\Docs\Financial
C:\Users\User1\Desktop\Docs\Financial\Collections Dispute
C:\Users\User1\Desktop\Docs\Job Seeking
C:\Users\User1\Desktop\Docs\Job Seeking\2020-2022 Resume Work
C:\Users\User1\Desktop\Docs\Job Seeking\2020-2022 Resume Work\Previous Resumes
C:\Users\User1\Desktop\Docs\Job Seeking\Archive
C:\Users\User1\Desktop\Docs\Job Seeking\Personal Info
C:\Users\User1\Desktop\Docs\OBS Profiles
C:\Users\User1\Desktop\Docs\OBS Profiles\1080pVR
C:\Users\User1\Desktop\Docs\Personal
C:\Users\User1\Desktop\Docs\Recipes
C:\Users\User1\Desktop\Docs\Vehicle-Related
C:\Users\User1\Desktop\Docs\Vehicle-Related\2011 Service Manual
C:\Users\User1\Desktop\Docs\Vehicle-Related\2021-08 Sunroof Claim
C:\Users\User1\Desktop\Docs\Vehicle-Related\Door Damage

.LINK
GitHub: https://github.com/jross365/Get-ExtendedFileAttributes

#>
function Get-Folders {
    [CmdletBinding()] 
    param( 
        [Parameter(Mandatory=$False)] [string]$Directory=((Get-Location).ProviderPath), 
        [Parameter(Mandatory=$False)] [switch]$SuppressErrors,
        [Parameter(Mandatory=$False)] [switch]$Recurse,
        [Parameter(Mandatory=$False)] [switch]$NoSort,
        [Parameter(Mandatory=$False)] [switch]$IgnoreExclusions,
        [Parameter(Mandatory=$False)] [switch]$IncludeRoot
        
    )

$Exclusions = 'filehistory|windows|recycle|@'
$Dirs = New-Object System.Collections.ArrayList

$EnumDirs = {[System.IO.Directory]::EnumerateDirectories("$Dir","*","TopDirectory")}



#region Validate parameters
If ($Directory.Length -eq 0 -or $Directory -eq $null){$Directory = (Get-Location).ProviderPath}
Else {If (!(Test-Path $Directory)){throw "$Directory is not a valid path"}}

If (!($IgnoreExclusions.IsPresent)){
    
    If ($Directory -match $Exclusions){throw "Path $Directory contains an excluded string. Please use the -IgnoreExclusions parameter"}

    $NotTheseNames = {$_ -inotmatch $Exclusions}

}

Else {$NotTheseNames = {$_ -ne $null}} #Have to put something here, or the .Where statements break

If ($IncludeRoot.IsPresent){$Dirs.Add($Directory) | out-null}

#endregion

$Dir = $Directory #Set for the "root" level enumeration

Try {(&$EnumDirs).Where($NotTheseNames).ForEach({$Dirs.Add($_) | Out-Null})}
catch {throw "Unable to enumerate directories in $Directory"}

If ($Recurse.IsPresent){
    
    $DirCount = $Dirs.Count

    $SubDirs = $Dirs
    
    Do {
    
        $DirQueue = New-Object System.Collections.ArrayList

        :DirLoop Foreach ($Dir in $SubDirs){
            
            Try {(&$EnumDirs).Where($NotTheseNames).ForEach({$DirQueue.Add($_) | Out-Null})}
            Catch {
                
                If (!$SuppressErrors.IsPresent){Write-Error "$($Error[0].Exception.Message.Split(':',2)[1].Trim())"}
                Continue DirLoop
            
            } #Close Catch
            
        } #Close ForEach $Dir
    
        $DirQueue.ForEach({$Dirs.Add($_) | Out-Null})

        $SubDirs = $DirQueue
    
        $DirCount = $SubDirs.Count
    
    } #Close Do
    
    Until ($DirCount -eq 0)

            } #Close If $Recurse.IsPresent

switch ($NoSort.IsPresent){

    $True {return $Dirs}

    $False {return ($Dirs | Sort-Object)}

} #CloseSwitch

        }

<#
.SYNOPSIS
Returns a file extension.

.DESCRIPTION
For a provided file path, returns the extension of the file (including '.')
Function doesn't care whether the path is valid or the file is real.

.PARAMETER FilePath
The path of the file to evaluate the extension of.

.INPUTS
None. You cannot pipe objects to Get-FileExtension.

.OUTPUTS
String. Get-FileExtension returns the file extension as a string.

.EXAMPLE
Get-FileExtension -FilePath ext.json

".json"

Note: Actual output does not contain quotes ("").

.LINK
GitHub: https://github.com/jross365/Get-ExtendedFileAttributes

#>
Function Get-FileExtension {
    param([string]$FilePath)

    return ([System.IO.Path]::GetExtension("$FilePath"))

} 

<#
.SYNOPSIS
Enumerates files.

.DESCRIPTION
Enumerates and returns all files in a provided directory.

.PARAMETER Directory
The directory to operate against.

.PARAMETER ExcludeFullPath
If specified, the file names will not include the file's directory path.

.PARAMETER Filter
Apply an inclusional filter to the output. Is a positional parameter (last position).
If specified, only paths/files matching the string or expression will be returned.

Note: Filter parameter does not respect asterisks.

.INPUTS
None. You cannot pipe objects to Get-Files.

.OUTPUTS
System.Collections.ArrayList. Get-Files returns an arraylist of string objects.

.EXAMPLE
Get-Files

C:\Users\User1\Desktop\Docs\recipes\Baked Chicken Wings - Seriously the BEST Crispy Baked Chicken Wings!.pdf
C:\Users\User1\Desktop\Docs\recipes\Chuck Roast Recipe.docx
C:\Users\User1\Desktop\Docs\recipes\Curry-Roux-chicken-curry-recipe.docx
C:\Users\User1\Desktop\Docs\recipes\Sweet Potato Pie Recipe.docx

.EXAMPLE
Get-Files -ExcludeFullPath

Baked Chicken Wings - Seriously the BEST Crispy Baked Chicken Wings!.pdf
Chuck Roast Recipe.docx
Curry-Roux-chicken-curry-recipe.docx
Sweet Potato Pie Recipe.docx

.EXAMPLE
Get-Files -ExcludeFullPath .pdf

Baked Chicken Wings - Seriously the BEST Crispy Baked Chicken Wings!.pdf

.LINK
GitHub: https://github.com/jross365/Get-ExtendedFileAttributes

#>
Function Get-Files {
param([string]$Directory,
[switch]$ExcludeFullPath,
[parameter(Position=2)][string]$Filter
)

If ($Directory.Length -eq 0 -or $null -eq $Directory){$Directory = (Get-Location).ProviderPath}

If ($null -eq $Filter -or $Filter.length -gt 0){
    $Filter = $Filter -replace '\*'
    $FilterBlock = {$_ -imatch "$Filter"}
    
    }
Else {$FilterBlock = $null}

Try {$Files = [System.IO.Directory]::EnumerateFiles("$Directory","*.*","TopDirectoryOnly")}
Catch {throw "$($_.Exception.Message)"}

$FilesNormalized = New-Object System.Collections.ArrayList
$Files.Where({$_ -notmatch 'thumbs.db'}).ForEach({$FilesNormalized.Add($_) | Out-Null})

If ($ExcludeFullPath.IsPresent -and $FilesNormalized.Count -ne 0){ 
    
    (0..($FilesNormalized.Count - 1 )).ForEach({
        
        $FilesNormalized[$_] = $FilesNormalized[$_].Replace("$Directory",'').TrimStart('\\')
    
    }) 

}

If ($null -ne $FilterBlock){return ($FilesNormalized.Where($FilterBlock))}
Else {Return $FilesNormalized}

}

Function Get-ExtendedAttributes {

[CmdletBinding()] 
param( 
    [Parameter(Mandatory=$False)] [string]$Path=((Get-Location).ProviderPath), 
    [Parameter(Mandatory=$False)] [switch]$Recurse,
    [Parameter(Mandatory=$False)] [switch]$WriteProgress,
    [Parameter(ParameterSetName='HelperFile',Mandatory=$False)] [switch]$UseHelperFile,
    [Parameter(ParameterSetName='HelperFile',Mandatory=$False)] [string]$HelperFilename="exthelper.json",
    [Parameter(Mandatory=$False)] [array]$Exclude,
    [Parameter(Mandatory=$False)] [array]$Include,
    [Parameter(Mandatory=$False)] [switch]$OmitEmptyFields,
    [Parameter(Mandatory=$False)] [switch]$ReportAccessErrors,
    [Parameter(Mandatory=$False)] [string]$ErrorOutFile
)

begin {

#region check variables
If (!(Test-Path -Path $Path)){throw "$Path is not a valid path"}

If ($UseHelperFile.IsPresent){

    Try {$JSON = Get-Content $HelperFilename -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop}
    Catch {throw "Helper file $HelperFilename is not valid"}

    If (($JSON[0].psobject.Properties.Name -join ',') -ne "Extension,Attrs"){throw "$HelperFilename does not contain the expected properties: Extension, Attrs"}
}

If ($ReportAccessErrors.IsPresent -and $ErrorOutFile.Length -gt 0){
    
    Switch ((Test-Path $ErrorOutFile)){

    $True {
            Try {"" | Out-File -Append $ErrorOutFile -ErrorAction Stop}
            Catch {throw "Error file $ErrorOutFile exists, but isn't writable"}
    }

    $False {

        Try {"" | Out-File $ErrorOutFile -ErrorAction Stop}
        Catch {throw "Error file $ErrorOutFile couldn't be created"}

    }

    } #Close Switch

    If ($ErrorOutFileError -eq $True){Write-Verbose "Could not append errors to $ErrorOutFile, writing to console:" -Verbose; $Exceptions.ForEach({Write-Error "$_"})}
 
} #Close if ErrorsToFile is present

#endregion

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
If ($Exclude.Count -gt 0){$OutFilter = $Exclude -join '|'; $OutFilterEnabled = $true}
Else {$OutFilterEnabled = $false}

If ($Include.Count -gt 0){$InFilter = $Include -join '|'; $InFilterEnabled = $true}
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

    $JSON.ForEach({
                    
                    [System.Collections.ArrayList]$HelperAttrs = $_.Attrs
                    $NoValueAttrs.ForEach({$HelperAttrs.Remove($_)})
                    $HelperHash.Add(($_.Extension),$HelperAttrs)
    
                })

    Remove-Variable JSON -ErrorAction SilentlyContinue
    &$ReclaimMemory

  #endregion

} #Close If UseHelper.IsPresent

#endregion

} #Close Begin

Process {

#region Build Directory/File Hashtable

$DirIndex = @{}
$Exceptions = New-Object System.Collections.ArrayList # To store

$DirIndex.Add("$Path",(Get-Files -Directory $Path -ExcludeFullPath))

If ($Recurse.IsPresent){
    
      $SubDirs = (Get-Folders -Directory $Path -SuppressErrors -Recurse)
          
      
      :DirLoop Foreach ($Dir in $SubDirs){

            Try {$SubdirFiles = Get-Files -Directory $Dir -ExcludeFullPath}
            Catch {$Exceptions.Add("$($Error[0].Exception.Message)") | Out-Null; Continue DirLoop}
            
            $DirIndex.Add("$Dir",$SubdirFiles)}
      
      Remove-Variable SubDirs -ErrorAction SilentlyContinue
      &$ReclaimMemory

} #Close If Recurse.IsPresent

$KeyDirs = $DirIndex.GetEnumerator().Name | Out-String -Stream

If ($OutFilterEnabled -eq $true){$KeyDirs = $KeyDirs.Where({$_ -inotmatch "$OutFilter"})}

<#
 Commented this out because "InFilter" works best against full file paths, and not against folders:
If ($InFilterEnabled -eq $true) {$KeyDirs = $KeyDirs.Where({$_ -imatch "$InFilter"})}
#>

#endregion Build Directory

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

          $FileAttrs = $HelperHash."$(Get-FileExtension -FilePath $File)"

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

Remove-Variable DirIndex,KeyDirs,Files,FileAttrs,TargetColumns,Object -ErrorAction SilentlyContinue
&$ReclaimMemory

#endregion

} #Close Process

end {
    
    #region Report errors, if requested:
        
    #Report Exceptions, if specified
    If ($ReportAccessErrors.IsPresent -and $Exceptions.Count -gt 0){
         
        switch (($ErrorOutFile.Length -gt 0)){
        
        $True {

                Try {$Exceptions | Out-File -Append $ErrorOutFile -ErrorAction Stop}
                Catch {
                        
                        Write-Verbose "Could not append errors to $ErrorOutFile, writing to console:" -Verbose
                        $Exceptions.ForEach({Write-Error "$_"})

                            } #Close Catch

        } #Close True

        $False {$Exceptions.ForEach({Write-Error "$_"})}

        } #Close switch
     
} #Close if $ReportAccessErrors.IsPresent

  #endregion
    
If ($OmitEmptyFields.IsPresent){
   
     #region Identify Unique Properties and preserve Property Order
        $PropertiesHash = @{}

        #The fastest way to total up the number of properties:
        $TotalProps = @{} 
        
        Foreach ($Line in $Results){

        ($Line.psobject.Properties.Name).ForEach({
            
            $Property = $_
        
            If ($null -eq ($TotalProps."$Property")){$TotalProps.Add("$Property",0)}

            })
       
        }
        
        $TotalPropsCount = $TotalProps.Count
        
        Remove-Variable TotalProps; &$ReclaimMemory

        $Results.ForEach({

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

        })

        $PropertyArrangement = New-Object System.Collections.ArrayList

        ($PropertiesHash.Keys).foreach({$PropertyArrangement.Add([pscustomobject]@{"Name" = $_; "Index" = $(($PropertiesHash.$_)[0])}) | Out-Null })

        $AllProperties = ($PropertyArrangement | Sort-Object "Index").Name

        $UsedProperties = New-Object System.Collections.ArrayList

        Remove-Variable PropertiesHash,TotalPropsCount,PropertyArrangement -ErrorAction SilentlyContinue
        &$ReclaimMemory

        $AllProperties.ForEach({

        $Property = $_

        If (($Results.Where({$_."$Property".Length -ne 0})).Count -gt 0){$UsedProperties.Add($Property) | Out-Null}

        })

        #endregion Identify Unique Properties

        $Results = $Results | Select-Object $UsedProperties

    } #Close if reduceDown is present

#endregion

return $Results

} #Close End

} #Close Function