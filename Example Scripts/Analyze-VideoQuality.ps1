#06/03: This is a work in progress:

param(
     [Parameter()]
     [string]$Path
 )

Try {import-module Get-ExtendedAttributes -ErrorAction Stop}
Catch {throw "Unable to import GEA module"}

If (!(Test-Path $Path)){throw "$Path is not a valid path"}

$VideoData = New-Object System.Collections.ArrayList

$VideoFields = @(
        "Name",
        @{N="FullName";E={$_.Path}},
        @{N="Extension"; E={$_.'File extension'}},
        "Size",
        "Length",
        @{N="FPS";E={($_.'Frame rate' -split ' ')[0]}},
        @{N="Height";E={$_.'Frame height'}},
        @{N="Width";E={$_.'Frame width'}},
        @{N="Audio Bitrate (kbps)"; E={$_.'Bit rate'}},
        @{N="Video Bitrate"; E={$_.'data rate'}},
        'Total bitrate'
        )

$geaParams = @{
    Path = "$Path"
    Recurse = $true
    WriteProgress = $true
    UseHelperFile = $true
    HelperFileName = "$HelperFile"
    Include = @(".avi",".mp4",".mpg",".mkv",".wmv")
    Clean = $true
    }

If ($null -eq $HelperFile -or $HelperFile.Length -eq 0){$geaParams.Remove("UseHelperFile")}

 #SPLAT!
Try {$VideoFiles = gea @geaParams}
catch {throw "The provided parameters aren't correct: $($geaParams.GetEnumerator() | Out-String)"}

$VideoFiles = $VideoFiles.Where({$_.Kind -eq 'Video'})
If ($VideoFiles.Count -eq 0){Write-Host "No files were found." -ForegroundColor Red; break}

#Remove those pesky, weird "tall arrow" characters:
$VideoFiles = ($VideoFiles | Convertto-csv -NoTypeInformation) -replace "$([char]8206)" | convertfrom-csv | Select-Object $VideoFields

#Preserve our field names, and get rid of the expressions
$VideoFields = $VideoFiles[0].psobject.Properties.Name

ForEach ($File in $VideoFiles){
    
    $Object = New-Object System.Object
    $VideoFields.ForEach({$Object | Add-Member -MemberType NoteProperty -Name "$_" -Value ($File.$_)})

    $Seconds = 0
    $Time = ($File.Length -split ':')
    $LengthSeconds = [system.math]::Round(([int]$Time[0] * 3600) + ([int]$Time[1]*60) + $Time[2])

    $FileSize = ($File.Size -split ' ')
    [int]$Size = $FileSize[0]
    $Unit = $FileSize[1]

    Switch ($Unit){

        {$_ -ieq "KB"}{$SizeKB = $Size}

        {$_ -ieq "MB"}{$SizeKB = $Size * 1024}

        {$_ -ieq "GB"}{$SizeKB = $Size * [system.math]::Pow(1024,2)}

        {$_ -ieq "TB"}{$SizeKB = $Size * [system.math]::Pow(1024,3)}

    }

    [int]$UncompressedSizeKB = [int]($File.'Total Bitrate' -replace 'kbps') * $LengthSeconds

    $CompressionDenominator = [System.Math]::Round($UncompressedSizeKB / $SizeKB,2)
    [int]$CompressionPercentage = (1 - (1 / $CompressionDenominator)) * 100

    $Object | Add-Member -MemberType Noteproperty -Name "Compression 1" -Value "$CompressionDenominator"
    $Object | Add-Member -MemberType Noteproperty -Name "Compressed (%)" -Value "$CompressionPercentage"
    $Object | Add-Member -MemberType NoteProperty -Name "Uncompressed"
}
