# Get-ExtendedAttributes

Get-ExtendedAttributes (**gea**) is a Powershell module for accessing the extended attributes of files.

## Description

The Get-ExtendedAttributes module provides functionality similar to the *Get-ChildItem* cmdlet. Instead of basic file attributes, however, **gea** enumerates and returns attributes not easily exposed to Powershell.

**These attributes include:**

* Video (*image/sound/combined bitrates, length, resolution, encoding, etc*)
* Music (*bitrate, length, artist, album, title, track, etc*)
* Image/EXIF (*resolution, camera information, focal length, ISO, orientation, etc*)
* Contacts (*Name, address, street, phone number, email address, etc*)
* Email (*To, From, Attachments, CC, BCC, send/received dates, subject, etc*)
* Documents (*Due date, word count, last-printed, last-saved, classification, pages, etc*)

* [And many more](https://stackoverflow.com/questions/22382010/what-options-are-available-for-shell32-folder-getdetailsof/62279888#62279888)

![Example of Photo management](/images/photosexample.png)


After completing, the function returns an [arraylist] of [pscustomobject] objects with the following properties:

* Hop:  Interval of the next hop in the transmission path
* RTT1: Round Trip Time of the first ICMP packet, in milliseconds (ms)
* RTT2: Round Trip Time of the second ICMP packet, in milliseconds (ms)
* RTT3: Round Trip Time of the third ICMP packet, in milliseconds (ms)
* Hostname: The DNS-resolved hostname of the hop's routing device
* IPAddress: The IP address of the hop's routing device

## Getting Started

### Dependencies

PS-TraceRoute.ps1 has the following dependencies:
* A Microsoft Windows operating system (Recommend Windows 7 or greater, Windows Server 2008 or greater)
* Windows Powershell (5+ recommended)


### Installing
This script can be copied and run from any location, and has no installation prerequisites or path requirements.

### Executing program

Running the script is simple:

* Import the function:
```
. .\PS-TraceRoute.ps1
```

* Run the function (example stores results in $RouteResults variable):
```
$RouteResults = Trace-Route -Destination <IPAddress>
```


## Help

I haven't encountered any problems with this function, more than likely due to its simplicity.

Because tracert.exe uses the standard Windows command line error handler, the cleanest and simplest way to check the success/fail status of tracert.exe is by looking at the $LASTEXITCODE variable.

The function contains an If statement to do exactly this:
```
If ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne $null){throw "traceroute returned LastExitCode $LASTEXITCODE"}
```

## Authors

I am the author.

## Version History

* 1.0 - Initial version.
    * It just works.

## License

This project is licensed under the GNUv3 License - see the LICENSE.md file for details.
