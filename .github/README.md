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

**Note:** The module exports the path of a "Helper File" as variable $HelperFile. This is used to *greatly* improve the speed and efficiency. Details on how to use the helper file are explained in the **Helper File** below.


### Using Get-ExtendedAttributes

* Running the function is simple:

```
Get-ExtendedAttributes
```

Alternatively, you can use the alias:

```
gea
```

## Parameters

**gea** contains many parameters to enhance its functionality and applicability.


### **Path**
The path of the directory or file you wish to retrieve extended attributes from.

This parameter is positional (*position 0*), and can be used without being named.

If unspecified, *-Path* uses the present working directory.


### **Recurse**
In cases where *-Path* is a directory, *-Recurse* will enumerate all subfolders and files within the provided path.

If *-Path* specifies a filename, *-Recurse* is ignored.


### **WriteProgress**
Displays a progress bar to support your mental health and welfare.

The progress bar reports which file it's enumerating attributes for, and displays the overall file progress.


### **UseHelperFile**
Instructs the function to use a Helper File.

Details about what a Helper File is and how to use it are written in the **Helper File** section below.

### **HelperFileName**
Provides the function with the path of the Helper File to use.

**Note:** *-UserHelperFile* and *-HelperFileName* will be consolidated into a single parameter in the future (*soon!*).


### **Exclude**
Applies an exclusionary ("*where not match*") filter on subfolders and files. If *-Path* is a file, *-Exclude* is ignored.

To specify more than one filter, comma-separate the strings you'd like to exclude.

This example excludes all files and folders containing ".png" or ".ps1" anywhere in the filename:

```
$N = Get-ExtendedAttributes -Exclude .png,.ps1
```
**Note:** *-Exclude* does not respect asterisks. If there's a desire to use asterisks for filtering, ask and I'll write the feature in. (*Or do it yourself, it's open source!*)


### **Include**
Applies an inclusionary ("*where match*") filter **for files only**. If *-Path* is a file, *-Include* is ignored.

As with *-Exclude*, you can comma-separate multiple strings you'd like to include. 

Also as with *-Exclude*, *-Include* does not respect asterisks.


### **OmitEmptyFields**
Instructs the function to remove all columns in the resultant data which do not contain any values. 

For example, a set of values that looks like this:

| Name  | Address | Street | Phone Number   | Email Address | 
| ----- |-------- | ------ | -------------- | ------------- |
| John  | 123 st  |        |                |               |
| Jake  |         |        |                | no@ip.org     |
|       | 3rd ave.|        |                | yep@nope.com  |

Would be reduced to these fields:

| Name  | Address | Email Address | 
| ----- |-------- | ------------- |
| John  | 123 st  |               |
| Jake  |         | no@ip.org     |
|       | 3rd ave.| yep@nope.com  |

**Note**: *-Clean* is an alias of *-OmitEmptyFields*.

This operation can take a lot of time, depending on how many files and specifie attributes reside in the dataset.

As with attribute lookups, the Helper File also reduces the number of possible empty fields. For this reason, it is *strongly* recommended that a Helper File be used when using *-OmitEmptyFields*.


### **ReportAccessErrors**
Reports all "Access Denied" errors to the console after the resultant data has been processed. Error reporting does not impact enumeration against files that were accessible.

### **ErrorOutFile**
Instructs the function to send errors to a designated text file instead of to the console. 


## The Helper File

**gea** is very quick when run without additional parameters for one or a small number of files.

In cases where there are a *large* number of files, the time it takes to query 500 attributes can cause **gea** to take a *very* long time to complete.

Thankfully, there's a clever solution to this problem.


### Understanding the Helper File

The Helper File is simply a JSON file called *exthelper.json*. It contains Keys (file extensions) and Values (applicable attributes for each file extension).

**gea** uses this file to limit the attribute retrievals to *only* attributes that are used by specific file types.

Instead of querying 500 attributes, when using the Helper File it will only query 30-40 (depending on the file type). This improves **gea**'s performance substantially.


### Where to get the Helper File

I have included a Helper File with the module that contains 315 extensions. This was generated from files on my systems and storage, and works perfectly (for me).


### How to use the Helper File

* If present when importing the module, the path to the Helper File is automatically assigned to the variable *$HelperFile*:

![Import-Module](/.github/images/import-module.png)

* When running **gea**, use the following parameters to use the Helper File:

```
Get-ExtendedAttributes -UseHelperFile -HelperFileName $HelperFile
```

**Note:** 👷 I realize how redundant it is to have a switch and an input variable for a single purpose. I will simplify this in the future 👷


### How much *does* the Helper File actually improve performance?

That's a fair question. Here's a test against 410 files:

![HelperFile Speed](/.github/images/helperspeed.png)


| Without Helper| With Helper   |
| ------------- | ------------- |
| 83 seconds    | 10 seconds    |
| 4.94 files/sec| 41 files/sec  |

That's a difference of **8 times** faster when using the Helper File!


### How to make your own Helper File

If you have a unique or specific set of file types that aren't included in the provided set, I have included a function to create your own Helper File (**New-AttrsHelperFile**)

Details on how to use this tool are written in the **Other Functions** section below.

# Other Functions
If **gea** is the star of the show, then there's also a supporting cast. Without them, the show wouldn't be possible.

This section briefly covers the other functions included in this Powershell module.

## **Reinventing the Wheel?**
When considering how to enumerate files and folders, I started with these three requirements.

1. Faster enumeration and simple/flat output of files and folders
2. Avoid enumerating attributes for the sake of efficiency
3.  Maintain all code within native Powershell/.NET Framework 

The first two requirements rule out *Get-ChildItem*, because **gci** is notoriously slow and enumerates attributes, which I was explicitly trying to avoid. And the last requirement rules out the cmd.exe "*dir.exe*" command, which is very fast but would require a wrapper function.

This left no "*off-the-shelf*" options (that I'm aware of), and I didn't want to borrow someone else's code. So I wrote **Get-Folders** and **Get-Files**.

## **Get-Folders**
**Get-Folders** is a function that enumerates directories in a provided path, and can do so recursively. **gfo** is an alias of **Get-Folders**.

### Using Get-Folders

* Running the function is simple:

```
Get-Folders
```

Alternatively, you can use the alias:

```
gfo
```

### Parameters
**gfo** contains a few parameters to enhance its functionality.

#### **Directory**
The path of the directory you wish to enumerate directories within.

If *-Directory* is not specified, **gfo** uses the present working directory.

#### **SuppressErrors**
Suppresses access errors without any output. This parameter can be useful when working with directory structures containing folders you know you don't have access to.

#### **Recurse**
Instructs the function to recursively enumerate all subdirectories within the specified path.

#### **NoSort**
*-NoSort* instructs the function to bypass alphabetical/hierarchical sorting of the enumerated directories, and returns the directories in the order they were discovered.

#### **IgnoreExclusions**
By default, **gfo** filters out directory paths matching the following strings:

* "filehistory"
* "windows"
* "recycle"
* "@"

Specifying *-IgnoreExclusions* will include directory paths with the strings listed above.

#### **IncludeRoot**
*-IncludeRoot* adds the "root" (*-Directory*) directory to the returned list of discovered directories/subdirectories.


## **Get-Files**
**Get-Files** is a function that enumerates files in a specified path.

**Note: Get-Files** does not operate recursively.

### Using Get-Files

* Running this function is also very simple:

```
Get-Files
```

Alternatively, you can use the alias:

```
gfi
```

### Parameters
**gfi** contains a handful of parameters to enhance its functionality.

#### **Directory**
The path of the directory you wish to enumerate the files within.

If not specified, **gfi** uses the present working directory.

#### **ExcludeFullPath**
By default, **gfi** provides the full directory path of every file it discovers. When specifying **-ExcludeFullPath**, only the file names are returned.

#### **Filter**
Applies an inclusional ("matches") filter to the output. *-Filter* is a positional parameter (last position), and doesn't require being named.

**Note:** *-Filter* only allows a single filter to be specified. If you'd like the multi-filter functionality (like **gea** provides), let me know and I'll include it.

Here's an example of how to only show files of type ".txt" using the *-Filter* parameter:

```
gfi -Filter *.txt
```

## **Get-FileExtension**
**Get-FileExtension** is a function that returns the file extension of a given file. 


### Using Get-Files

* Running this function requires specifying the filename:

```
Get-FileExtension -FilePath D:\somefile.txt

.txt
```

However, the *actual* path doesn't matter. You can pass the function nonsense, and it'll return the perceived file extension:
```
Get-FileExtension asdfq234r3e2f.sql

.sql
```
### Parameters
**Get-FileExtension** has a single parameter:

#### **FilePath**
The name or path of the file.


## **New-AttrsHelperFile**
**New-AttrsHelperFile** is a function that analyzes CSV files with contents created by **gea** to generate a new Helper File.

You may want to create a new helper file if your use-case for this module applies to files whose extensions aren't included in the provided *extHelper.json* file.

### Notes and Recommendations
These are some general thoughts on the best way to use **New-AttrsHelperFile**:

* When running **gea** to generate the initial data, it's best to specify *-OmitEmptyFields* upfront. This will save a lot of time when re-analyzing the data.
* When saving the **gea** output as a CSV, make sure you specify the **Export-Csv* cmdlet's *-NoTypeInformation* parameter.
* You don't need a huge dataset of files in each CSV to generate a perfect *exthelper.json* file. 
    * Consider hand-picking a small quantity of each file type that you think will have the desired properties.
* 


### Using Get-Files

* Running this function requires specifying the filename:

```
Get-FileExtension -FilePath D:\somefile.txt

.txt
```

However, the *actual* path doesn't matter. You can pass the function nonsense, and it'll return the perceived file extension:
```
Get-FileExtension asdfq234r3e2f.sql

.sql
```
### Parameters
**Get-FileExtension** has a single parameter:

#### **FilePath**
The name or path of the file.



## Help

I haven't encountered any problems with this function, more than likely due to its simplicity.

Because tracert.exe uses the standard Windows command line error handler, the cleanest and simplest way to check the success/fail status of tracert.exe is by looking at the $LASTEXITCODE variable.

The function contains an If statement to do exactly this:
```
If ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne $null){throw "traceroute returned LastExitCode $LASTEXITCODE"}
```


### 👷 Work In Progress!
Everything below this point is in the process of being written. Please check in periodically for updates as this documentation is created and completed. (*Last Update: 05/31/2022*)

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
