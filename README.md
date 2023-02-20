# ZoneStripper v1.2
Removes the Zone.Identifier alternate data stream that identifies files as 'from the internet'

![Screenshot](https://user-images.githubusercontent.com/7834493/220038927-b286b17e-25c3-43c4-a890-c0bd1c581d03.png)

With Microsoft taking away the option to click through warnings about macro enabled documents and load them anyway, it's becoming more important to 'unblock' these documents, among various other reasons you'd want to do this for other files. It's easy enough to do this manually for a single file through Explorer (however, this only changes, not removes, the zone identifier), but it might get tedious if you have a lot of files. ZoneStripper will recursively (or single level) go through a folder and completely remove the zone identifier from all files, making them just like any other file that came from your own computer rather than the internet.

Ironically, this requires no special permissions. ZoneStripper doesn't need admin permissions, and can remove it from any file you have read/write permission for. The attribute can't be removed from read-only file, so there's an option for how to handle this: (1) Skip read only files, (2) Clear the read only attribute, remove the zone identifier, and put the read only attribute back, or (3) Clear the read only attribute and leave it that way after removing the zone identifier.

This is based on an import of the VB6 version originally posted as a years-later update to an original demo for reading/writing them here:

[[VB6] Code Snippet: Get/set/del file zone identifier (Run file from internet? source)](https://www.vbforums.com/showthread.php?804967-VB6-Code-Snippet-Get-set-del-file-zone-identifier-(Run-file-from-internet-source))

New features to control what files to apply it to and whether to change the zone instead of remove entirely have been added, and it's been updated to use tbShellLib instead of oleexp and to support compiling for x64 (not much work here, just had to change 3 Longs to LongPtr). 

### Updates
**Version 1.2:** Minor bug fixes: Error in tabstop order and inability to set 'Untrusted' zone (if, for whatever reason, anyone ever wanted to).

### Requirements

-Source requires [twinBASIC Beta 239 or newer](https://github.com/twinbasic/twinbasic/releases) to open/build.

-As far as I know, this is only applicable to NTFS file systems. This project does use the documented COM interfaces for this rather than manually reading/writing the alternate data stream, so if Windows ever does support this on other file systems, there's still a good chance it works.

### How it works

NTFS supports [alternate data streams](https://www.malwarebytes.com/blog/news/2015/07/introduction-to-alternate-data-streams). These are hidden data blocks attached to a file that don't count towards it's size, so you don't even know if they're there without special utitilities. For instance, all major web browsers add this to all downloads: `C:\download\file.docxm:Zone.Identifier`, a data block containing a value indicating what security zone the file belongs to. This is how Windows knows if a file is from the internet. You can access that stream manually via VB's `Open` syntax, but it's easier, for this purpose at least, to use Window's built in handling:

```
Public Function GetFileSecurityZone(sFile As String) As URLZONE
'returns the Zone Identifier of a file, using IZoneIdentifier
'This could also be done by ready the Zone.Identifier alternate
'data stream directly; readfile C:\file.txt:Zone.Identifier

Dim lz As Long
Dim pZI As PersistentZoneIdentifier
Set pZI = New PersistentZoneIdentifier

Dim pIPF As IPersistFile
Set pIPF = pZI

pIPF.Load sFile, STGM_READ
pZI.GetId lz
GetFileSecurityZone = lz

Set pIPF = Nothing
Set pZI = Nothing

End Function

Public Sub SetFileSecurityZone(sFile As String, nZone As URLZONE)
'As suggested in the enum, you technically can set it to custom values
'If you do, they should be between 1000 and 10000.
Dim pZI As PersistentZoneIdentifier
Set pZI = New PersistentZoneIdentifier

pZI.SetId nZone
Dim pIPF As IPersistFile
Set pIPF = pZI
pIPF.Save sFile, 1

Set pIPF = Nothing
Set pZI = Nothing

End Sub

Public Sub RemoveFileSecurityZone(sFile As String)
Dim pZI As PersistentZoneIdentifier
Set pZI = New PersistentZoneIdentifier

pZI.Remove
Dim pIPF As IPersistFile
Set pIPF = pZI
pIPF.Save sFile, 1

Set pIPF = Nothing
Set pZI = Nothing
End Sub
```


The available zones are identifie by the following enum:


```
Public Enum URLZONE
    URLZONE_INVALID = -1
    URLZONE_PREDEFINED_MIN = 0
    URLZONE_LOCAL_MACHINE = 0
    URLZONE_INTRANET
    URLZONE_TRUSTED
    URLZONE_INTERNET
    URLZONE_UNTRUSTED
End Enum
```

Files are marked as `URLZONE_INTERNET` by web browsers. 

