# RegexCOM

This is my first attempt a creating a COM visible DLL. The intent
is to expose .NET Framework Regex methods in MS Office applications
via VBA.

Note that Microsoft recently implement a new RegExp class in VBA. 
This new class beasically replicates to functionality of the 
VBScript library's RegExp class. For most users this will probably
be adequate, but for users looking for features that exist in the
.NET Framework implementation but not the VBScript implementation,
this DLL is an option.
