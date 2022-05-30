# Get-ExtendedFileAttributes

A set of functions to efficiently enumerate extended file attributes

## Description

The function contained in the PS-TraceRoute.ps1 Powershell script parses traceroute (tracert.exe) output. 

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
