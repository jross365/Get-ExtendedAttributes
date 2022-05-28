
    #WorkBench.ps1 is for cleaning up/testing changes to functions and blocks of code
    
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
    
function Get-Folders ($Directory,[switch]$SuppressErrors,[switch]$Recurse,[switch]$NoSort,[switch]$IgnoreExclusions){

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