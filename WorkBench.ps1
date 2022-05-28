
    #WorkBench.ps1 is for cleaning up/testing changes to functions and blocks of code
    
   
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

#region Define preliminary variables
$Exclusions = 'filehistory|windows|recycle|@'
$Dirs = New-Object System.Collections.ArrayList

$EnumDirs = {[System.IO.Directory]::EnumerateDirectories("$Dir","*","TopDirectory")}

#endregion

#region Validate parameters
If ($Directory.Length -eq 0){$Directory = (Get-Location).ProviderPath}
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

        Function Get-Files ($Directory,[switch]$ExcludeFullPath){

            Try {$Files = [System.IO.Directory]::EnumerateFiles("$Directory","*.*","TopDirectoryOnly")}
            Catch {throw "$($_.Exception.Message)"} #Changed this because we're not leaning on recursion in this case
            
            $FilesNormalized = New-Object System.Collections.ArrayList
            $Files.Where({$_ -notmatch 'thumbs.db'}).ForEach({$FilesNormalized.Add($_) | Out-Null})
            
            If ($ExcludeFullPath.IsPresent -and $FilesNormalized.Count -ne 0){ 
                
                (0..($FilesNormalized.Count - 1 )).ForEach({
                    
                    $FilesNormalized[$_] = $FilesNormalized[$_].Replace("$Directory",'').TrimStart('\\')
                
                }) 
            
            }
            
            Return $FilesNormalized
            
            }