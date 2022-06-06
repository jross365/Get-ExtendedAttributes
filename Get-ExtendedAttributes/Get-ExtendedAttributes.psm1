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

.PARAMETER IgnoreExclusions
If specified, will include subdirectory search in FileHistory, Windows, Recycle Bin and '@' directories.

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
C:\Users\User1\Desktop\Docs\Job Seeking\Archive
C:\Users\User1\Desktop\Docs\Job Seeking\Personal Info
C:\Users\User1\Desktop\Docs\OBS Profiles
C:\Users\User1\Desktop\Docs\OBS Profiles\1080pVR
C:\Users\User1\Desktop\Docs\Personal
C:\Users\User1\Desktop\Docs\Recipes
C:\Users\User1\Desktop\Docs\Vehicle-Related
C:\Users\User1\Desktop\Docs\Vehicle-Related\2011 Service Manual
C:\Users\User1\Desktop\Docs\Vehicle-Related\Door Damage

.LINK
GitHub: https://github.com/jross365/Get-ExtendedAttributes

#>
function Get-Folders {
    [CmdletBinding()] 
    param( 
        [Parameter(Mandatory=$False)][Alias('d')] [string]$Directory=((Get-Location).ProviderPath), 
        [Parameter(Mandatory=$False)][Alias('q')] [switch]$SuppressErrors,
        [Parameter(Mandatory=$False)][Alias('r')] [switch]$Recurse,
        [Parameter(Mandatory=$False)] [switch]$NoSort,
        [Parameter(Mandatory=$False)][Alias('a')] [switch]$IgnoreExclusions,
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

If ($IncludeRoot.IsPresent){$Dirs.Add($Directory) | out-null}

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
GitHub: https://github.com/jross365/Get-ExtendedAttributes

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
GitHub: https://github.com/jross365/Get-ExtendedAttributes

#>
Function Get-Files {
param([Alias('d')][string]$Directory,
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

<#
.SYNOPSIS
Enumerates files and returns extended attributes.

.DESCRIPTION
Enumerates extended attributes (access to media, document, etc) for files in a similar format to "Get-ChildItem".

.PARAMETER Path
The directory to operate against.

.PARAMETER Recurse
If specified, recursively enumerate files in all subdirectories nested in the root directory.

.PARAMETER WriteProgress
If specified, will show the progress of file attributes enumeration with a progress bar.

.PARAMETER UseHelperFile
If specified, instructs function to use a "helper file" to speed up attributes enumeration.

.PARAMETER HelperFileName
If "-UseHelperFile" is specified, this parameter is mandatory. The full path of the helper file (JSON).

.PARAMETER Exclude
If specified, applies an exclusionary filter to files or folders. Comma-separate strings for multiple exclusions.

.PARAMETER Include
If specified, applies an inclusionary filter to files. Comma-separate strings for multiple exclusions.
Inclusionary filters aren't applied to folders, by design. 

.PARAMETER OmitEmptyFields
If specified, returned file attributes are filtered to include only attributes that aren't empty. Columns without any values are omitted.

.PARAMETER ReportAccessErrors
If specified, access errors encountered during enumeration ("Access denied") will be reported in the console after enumeration is complete.

.PARAMETER ErrorOutFile
If "-ReportAccessErrors" is specified, this parameter is optional. Writes encountered access errors to the specified file name.

.INPUTS
None. You cannot pipe objects to Get-ExtendedAttributes.

.OUTPUTS
System.Collections.ArrayList. Get-ExtendedAttributes returns an arraylist of string objects.

.EXAMPLE
Write-Host "Due to the complexity of this function and the possible use-cases and outputs, please see the GitHub documentation for examples.""

.LINK
GitHub: https://github.com/jross365/Get-ExtendedAttributes

#>
Function Get-ExtendedAttributes {

    [CmdletBinding()] 
    param( 
        [Parameter(Mandatory=$False,Position=0)] [string]$Path=((Get-Location).ProviderPath), 
        [Parameter(Mandatory=$False)] [switch]$Recurse,
        [Parameter(Mandatory=$False)] [switch]$WriteProgress,
        [Parameter(Mandatory=$False)][Alias('HF')] [string]$HelperFile,
        [Parameter(Mandatory=$False)] [array]$Exclude,
        [Parameter(Mandatory=$False)] [array]$Include,
        [Parameter(Mandatory=$False)][Alias('Clean')] [switch]$OmitEmptyFields,
        [Parameter(Mandatory=$False)] [switch]$ReportAccessErrors,
        [Parameter(Mandatory=$False)] [string]$ErrorOutFile
    )
    
    begin {
    
    #region check variables
    If (!(Test-Path -Path $Path)){throw "$Path is not a valid path"}
    
    $FSOInfo = Get-Item $Path
    $Path = $FSOInfo.FullName #Protection against relative paths

    If (($FSOInfo).Attributes -match 'Directory'){$FSOType = "FSO-Directory"; $NameSpace = $Path}
    Else {$FSOType = "FSO-File"; $NameSpace = $FSOInfo.Directory.FullName}

    If ($HelperFile.Length -gt 0){ write-verbose "Helper File detected" -Verbose
    
        Try {$JSON = Get-Content $HelperFile -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop}
        Catch {throw "Helper file $HelperFile is not valid"}
    
        If (($JSON[0].psobject.Properties.Name -join ',') -ne "Extension,Attrs"){throw "$HelperFile does not contain the expected properties: Extension, Attrs"}

        $UseHelperFile = $true
    }
    Else {$UseHelperFile -eq $false}


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
    Try {$RootPath = $ShellObj.Namespace($NameSpace)}
    Catch {throw "Unable to initialize Shell Application for namespace $Namespace"}
    
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
    If ($UseHelperFile -eq $true){
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
    $Exceptions = New-Object System.Collections.ArrayList # To store error output
    
    Switch ($FSOType){ #Actions to take based on source type (file or directory)

    {$_ -eq "FSO-Directory"}{
    
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

                            }

    {$_ -eq "FSO-File"}{
                        $DirIndex.Add("$($FSOInfo.Directory.FullName)","$($FSOInfo.Name)")
                        }

    } #Close Switch

    $KeyDirs = $DirIndex.GetEnumerator().Name | Out-String -Stream
        
    If (($OutFilterEnabled -eq $true) -and ($FSOType -eq "FSO-Directory")){$KeyDirs = $KeyDirs.Where({$_ -inotmatch "$OutFilter"})}
    #f (($InFilterEnabled -eq $true) -and ($FSOType -eq "FSO-Directory")) {$KeyDirs = $KeyDirs.Where({$_ -imatch "$InFilter"})}
   
    #endregion
    
    #region Filter hash table file values
    $KeyDirs.ForEach({
            $Dir = $_
        
            If (($OutFilterEnabled -eq $true) -and ($FSOType -eq "FSO-Directory")){$DirIndex."$Dir" = $DirIndex."$Dir".Where({$_ -inotmatch "$OutFilter"})}
            If (($InFilterEnabled -eq $true) -and ($FSOType -eq "FSO-Directory")){$DirIndex."$Dir" = $DirIndex."$Dir".Where({$_ -imatch "$InFilter"})}
    
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
            
                Switch ($UseHelperFile){
            
                $True {
            
                        $FileAttrs = $HelperHash."$(Get-FileExtension -FilePath $File)"
                
                        If ($FileAttrs.Count -gt 0){$TargetColumns = $FileAttrs}
                        Else {$TargetColumns = $ValidColumns}
                
                        }#Close True
                    
                $False {$TargetColumns = $ValidColumns}

                Default {$TargetColumns = $ValidColumns}
            
                }#Close Switch UserHelperFile.IsPresent
            
                $TargetColumns.ForEach({ 
                    
                    Try {$Object | Add-Member -ErrorAction Stop -Membertype Noteproperty -Name "$($AttrHash.$_)" -Value ($FolderObj.GetDetailsOf($FileObj, $_))}
                    Catch {$Object | Add-Member -ErrorAction Stop -Membertype Noteproperty -Name "$($AttrHash.$_)" -Value "-"; Continue}
                
                    })
            
                $Results.Add($Object) | Out-Null
            
                $x++ 
        
            }) #Close Files.ForEach
    
    } #Close :KeyLoop
    
    Remove-Variable DirIndex,KeyDirs,Files,TargetColumns,Object -ErrorAction SilentlyContinue
    &$ReclaimMemory
    
    #endregion
    
    } #Close Process
    
    end {
        
        #region Report errors, if requested:
            
        #Report Exceptions, if specified
        If ($ReportAccessErrors.IsPresent -and $Exceptions.Count -gt 0){
             
            Switch (($ErrorOutFile.Length -gt 0)){
            
            $True {
    
                    Try {$Exceptions | Out-File -Append $ErrorOutFile -ErrorAction Stop}
                    Catch {
                            
                            Write-Verbose "Could not append errors to $ErrorOutFile, writing to console:" -Verbose
                            $Exceptions.ForEach({Write-Error "$_"})
    
                                } #Close Catch
    
                    }
    
            $False {$Exceptions.ForEach({Write-Error "$_"})}
    
            } #Close switch
         
    } #Close if $ReportAccessErrors.IsPresent
    
      #endregion
        
    If ($OmitEmptyFields.IsPresent){
        
        #region Identify Unique Properties and preserve Property Order
        Switch ($FSOType){
        
        {$_ -eq "FSO-Directory"}{ 
            #If it's a directory, we have a lot of work to do:

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
    
            :LineLoop Foreach ($Line in $Results){
        
                $ObjProperties = $Line.psobject.Properties.Name
        
                Foreach ($Index in (0..$ObjProperties.GetUpperBound(0))){
        
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

                    If ($PropertiesHash.Count -eq $TotalPropsCount){break LineLoop} #As soon as we've encountered every possible property in $Results, break
        
                } 
        
            } #Close Foreach Line in $Results
        
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
        
            }

            {($_ -eq "FSO-File") -or ($Results.Count -eq 1)}{  #If it's a single file, our work is much simpler
        
                $UsedProperties = New-Object System.Collections.ArrayList
                
                $Results[0].psobject.Properties.Name.ForEach({If ($Results[0].$_.Length -gt 0){$UsedProperties.Add($_) | out-null}})
            
            }

        } #Close Switch

            #endregion Identify Unique Properties
    
            $Results = $Results | Select-Object $UsedProperties
    
        } #Close if reduceDown is present
    
    #endregion
    
    return $Results
    
    } #Close End
    
}

<#
.SYNOPSIS
Creates a helper file for use with Get-ExtendedAttributes.

.DESCRIPTION
Takes a folder with CSVs created by "Get-ExtendedAttributes" and analyzes the CSVs' contents to generate a new JSON helper file.

.PARAMETER Folder
The directory to operate against.

.PARAMETER SaveAs
The file name to save the resultant helper file as.

.PARAMETER WriteProgress
If specified, will show the the analysis progress as a progress bar.

.INPUTS
None. You cannot pipe objects to New-AttrsHelperFile.

.OUTPUTS
System.Collections.ArrayList. Get-ExtendedAttributes returns an arraylist of string objects.

.EXAMPLE
Write-Host "Due to the complexity of this function and the possible use-cases and outputs, please see the GitHub documentation for examples.""

.LINK
GitHub: https://github.com/jross365/Get-ExtendedAttributes

#>
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

#region Module Instructions
New-Alias -Name gea -Value Get-ExtendedAttributes
New-Alias -Name gfo -Value Get-Folders
New-Alias -Name gfi -Value Get-Files

Export-ModuleMember -Function Get-Folders
Export-ModuleMember -Function Get-FileExtension
Export-ModuleMember -Function Get-Files
Export-ModuleMember -Function Get-ExtendedAttributes
Export-ModuleMember -Function New-AttrsHelperFile

Export-ModuleMember -Alias gea
Export-ModuleMember -Alias gfo
Export-ModuleMember -Alias gfi

[system.collections.arraylist]$SplitPath = ((Get-Command Get-Folders).Module).Path -split '\\'
$SplitPath[$SplitPath.Count -1] = "exthelper.json"
$HelperFile = $SplitPath -join '\'
Remove-Variable SplitPath -ErrorAction SilentlyContinue

If (Test-Path $HelperFile){
    
    Export-ModuleMember -Variable HelperFile
    Write-Verbose 'Helper File found, path stored as variable $HelperFile' -Verbose
    Write-Verbose "Helper File location: $HelperFile" -Verbose

}

#endregion
