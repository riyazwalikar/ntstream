# NTStream - Search and manage Alternate Data Streams on NTFS partitions

## Introduction:
Alternate Data Streams (ADS) are a feature of the NTFS file system that allow files to have multiple data streams. For example, every file on a NTFS partition will have a primary data stream which is the file's actual content. This is accessible via `filename.txt:$DATA:""`

We can create additional data streams to this file and they would exist as `filename.txt:$DATA:file2.txt` and `filename.txt:$DATA:file3.txt` etc. ADS can also be created on folders apart from standard files.

## What is NTStream?
NTStream is small GUI application that allows you to rapidly search through your filesystem to find files that have an ADS and allows you to interact with it. This program uses the Windows backup API to read and list the ADS. Interaction with the discovered ADS is then carried out using standard Windows IO APIs.

You can search a particular file, a folder or subfolders using the GUI interface. Once a file with ADS is found, you can manipulate it, delete it or simply add new streams.

#### Create New Stream:
After you scan a folder or file, you can right click on any item in the list and select Create new Stream.
In the New Stream Creator wizard, browse for a file that you want to attach as ADS.
Type a name for your ADS or keep the default name.
Click on Create when done.


#### Viewing and Editing a Stream:
After a search, double click on an item that has ADS listed.
Under the Streams window that opens up, right click on the stream that you want to view and click View Stream.
The Stream Viewer will open up which will correctly display the contents of a file if they are text based.
You can even edit the stream and save it by clicking on Edit and then Save. Please do not edit binaries or other file formats.
 

#### Export Stream:
Double Click on an item that has ADS listed.
Under the Streams window that opens up, right click on the stream you want to export and click on Extract Stream.
Type a file name, with the correct extension if you what it is, and click Save.
The file will be extracted without any confirmations.


#### Delete Stream:
After a search, double click on an item that has ADS listed.
Under the Streams window that opens up, right click on the stream that you want to delete and click Delete Stream.
Click on yes at the confirmation message box.

## To Build
This was written in VB6 on an older version of Windows. Compiled binaries are available at request. Send me an email at riyazwalikar at gmail dot com.

If you want to build or play around with this code yourself, you will need:
- Visual Basic 6
- Register the components in \COM

## Known bugs
- The ADS delete option may not work if you do not have permission on the file. There is no error handling in that function yet :(
- Files larger than 64K may fail to be attached to another file as ADS owing to lack of threading while file is being read. This could be an enhancement in a new revision. 
