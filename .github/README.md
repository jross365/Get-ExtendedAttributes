# Get-ExtendedAttributes

Get-ExtendedAttributes (**gea**) is a Powershell module for accessing the extended attributes of files.

(👷This documentation is an active work in progress👷)

## Description

The Get-ExtendedAttributes module provides functionality similar to the *Get-ChildItem* cmdlet. Instead of basic file attributes, however, **gea** enumerates and returns attributes not easily exposed to Powershell.

**These attributes include:**

* **Video** (image/sound/combined bitrates, length, resolution, encoding, etc)
* **Music** (bitrate, length, artist, album, title, track, etc)
* **Image/EXIF** (resolution, camera information, focal length, ISO, orientation, etc)
* **Contacts** (Name, address, street, phone number, email address, etc)
* **Email** (To, From, Attachments, CC, BCC, send/received dates, subject, etc)
* **Documents** (Due date, word count, last-printed, last-saved, classification, pages, etc)

* [And many more](https://stackoverflow.com/questions/22382010/what-options-are-available-for-shell32-folder-getdetailsof/62279888#62279888)

![Example of Photo management](/.github/images/photosexample.png)


## Getting Started

### Dependencies

**Get-ExtendedAttributes** has the following dependencies:
* Microsoft Windows OS (Windows 7 or greater, Windows Server 2008 or greater)
* Windows Powershell (5.1+ recommended)


### Installing
Save the *Get-ExtendedAttributes* folder to any one of these three folders:

```
C:\Users\<username>\Documents\WindowsPowerShell\Modules
C:\Program Files\WindowsPowerShell\Modules        
C:\Windows\system32\WindowsPowerShell\v1.0\Modules
```

### Importing the Module

Import the module using the following command:

```
Import-Module Get-ExtendedAttributes
```

![Import-Module](/.github/images/import-module.png)

* **Note:** The module exports the path of a "Helper File" as variable $HelperFile. This is used to *greatly* improve the speed and efficiency of **gea**. Details on how to use the helper file are explained below.

### 👷 Work In Progress! 👷
Everything below this point is in the process of being written. Please check in periodically for updates as this documentation is created and completed. (*Last Update: 05/30/2022*)

### Using Get-ExtendedAttributes

* Running the script is simple:

```
Get-ExtendedAttributes
```
![Get-ExtendedAttributes](/.github/images/get-extendedattributes.png)

Alternatively, you can use the alias **gea**

```
gea
```
![gea](/.github/images/gea.png)


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

## Known Issues
* Strange/faulty behavior when working with files in UserProfile diectories (caused by NTUSER.DAT)
* Some file attribute values obtained from internet or non-Windows sources contain odd/wrong characters
    * I'm not convinced this is a problem with the module, but a future "helper" function may help fix these instances.

## License

This project is licensed under the GNUv3 License - see the LICENSE.md file for details.
