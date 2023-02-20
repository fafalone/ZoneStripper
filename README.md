# ZoneStripper
Removes the Zone.Identifier alternate data stream that identifies files as 'from the internet'

![Screenshot](https://user-images.githubusercontent.com/7834493/220021971-2111c9b8-60d5-44a4-840e-8070610e3990.jpg)

With Microsoft taking away the option to click through warnings about macro enabled documents and load them anyway, it's becoming more important to 'unblock' these documents, among various other reasons you'd want to do this for other files. It's easy enough to do this manually for a single file through Explorer (however, this only changes, not removes, the zone identifier), but it might get tedious if you have a lot of files. ZoneStripper will recursively (or single level) go through a folder and completely remove the zone identifier from all files, making them just like any other file that came from your own computer rather than the internet.

Ironically, this requires no special permissions. ZoneStripper doesn't need admin permissions, and can remove it from any file you have read/write permission for. The attribute can't be removed from read-only file, so there's an option for how to handle this: (1) Skip read only files, (2) Clear the read only attribute, remove the zone identifier, and put the read only attribute back, or (3) Clear the read only attribute and leave it that way after removing the zone identifier.

This is an import of the VB6 version originally posted as a years-later update to an original demo for reading/writing them here:

[[VB6] Code Snippet: Get/set/del file zone identifier (Run file from internet? source)]([VB6] Code Snippet: Get/set/del file zone identifier (Run file from internet? source))

It's been updated to use tbShellLib instead of oleexp and to support compiling for x64 (not much work here, just had to change 3 Longs to LongPtr). 

###Requirements

-Source requires [twinBASIC Beta 239 or newer](https://github.com/twinbasic/twinbasic/releases) to open/build.
