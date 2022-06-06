
<#
This example script enumerates and parses video quality information, using the Get-ExtendedAttributes module.

Parameters:
-Path:              The directory containing video files
-Recurse:           Indicates whether to recursively search the provided directory for video files
-Gridview:          Also present the resultant data using "Out-GridView" when finished
-CompressionInfo:   Calculate compression info in addition to returning file attributes

#>

param(
     [Parameter()]
     [string]$Path,
     [switch]$Recurse,
     [switch]$Gridview,
     [switch]$CompressionInfo
 )

Try {import-module Get-ExtendedAttributes -ErrorAction Stop}
Catch {throw "Unable to import gea module"}

If (!(Test-Path $Path)){throw "$Path is not a valid path"}

$VideoData = New-Object System.Collections.ArrayList

$VideoFields = @(
        "Name",
        @{N="FullName";E={$_.Path}},
        @{N="Extension"; E={$_.'File extension'}},
        "Size",
        "Length",
        @{N="FPS";E={[double]($_.'Frame rate' -split ' ')[0]}},
        @{N="Height";E={[int]($_.'Frame height')}},
        @{N="Width";E={[int]($_.'Frame width')}},
        @{N="Audio Bitrate (kbps)"; E={[int]($_.'Bit rate' -replace 'kbps')}},
        @{N="Video Bitrate (kbps)"; E={[int]($_.'Data rate' -replace 'kbps')}},
        @{N="Total Bitrate (kbps)"; E={[int]($_.'Total bitrate' -replace 'kbps')}}
        )

$geaParams = @{
    Path = "$Path"
    Recurse = ($Recurse.IsPresent)
    WriteProgress = $true
    UseHelperFile = $true
    HelperFileName = "$HelperFile"
    Include = @(".avi",".mp4",".mpg",".mkv",".wmv")
    Clean = $true
    }

 #SPLAT!
Try {$VideoFiles = (gea @geaParams).Where({$_.Kind -eq 'Video'})}
catch {throw "The provided parameters aren't correct: $($geaParams.GetEnumerator() | Out-String)"}

#Preserve our field names, and get rid of the expressions
$VideoFields = $VideoFiles[0].psobject.Properties.Name

#If we want to calculate compression information:
If ($CompressionInfo.IsPresent){

ForEach ($File in $VideoFiles){
    
    $Object = New-Object System.Object
    $VideoFields.ForEach({$Object | Add-Member -MemberType NoteProperty -Name "$_" -Value ($File.$_)})

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

    #region calculate overall compression

    [int]$UncompressedSizeKB = ($File.'Total Bitrate (kbps)' / 8) * $LengthSeconds

    $CompressionDenominator = [System.Math]::Round($UncompressedSizeKB / $SizeKB,2)
    [int]$CompressionPercentage = (1 - (1 / $CompressionDenominator)) * 100

    #endregion

    #region calculate Audio:Video ratio

    $AudioPerc = ($Object.'Audio Bitrate (kbps)' / $Object.'Total Bitrate (kbps)')
    $VideoPerc = ($Object.'Video Bitrate (kbps)' / $Object.'Total Bitrate (kbps)')

    $AVRatio = [system.math]::Round(($VideoPerc / $AudioPerc),2)

    #endregion

    #region calculate compressed/uncompressed data density from resolution, framerate and data bitrate

    $PixelDensity = $File.Height * $File.Width

    $PixelsPerSec = $PixelDensity * $File.FPS

    $DataBitRate = $File.'Video Bitrate (kbps)' * 1024

    $UnCompDataDensity = [system.math]::Round((($DataBitRate / $PixelsPerSec) * 100),2)

    $CompDataDensity = (1 - ($CompressionPercentage / 100)) * $UnCompDataDensity

    #endregion

    $Object | Add-Member -MemberType Noteproperty -Name "AV Ratio (1:)" -Value $AVRatio
    $Object | Add-Member -MemberType NoteProperty -Name "RawSize(MB)" -Value ([int]($UncompressedSizeKB / 1024))
    $Object | Add-Member -MemberType NoteProperty -Name "RawDensity" -Value $UnCompDataDensity
    $Object | Add-Member -MemberType Noteproperty -Name "CompRatio (1:)" -Value "$CompressionDenominator"
    $Object | Add-Member -MemberType Noteproperty -Name "Compressed (%)" -Value "$CompressionPercentage"
    $Object | Add-Member -MemberType NoteProperty -Name "CompDensity"    -Value ([system.Math]::Round($CompDataDensity,2))
    

    $VideoData.Add($Object) | Out-Null

}

}

Else {$VideoData = $VideoFiles}

If ($Gridview.IsPresent){$VideoData | Out-GridView}

Return $VideoData
