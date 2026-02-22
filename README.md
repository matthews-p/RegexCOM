# RegexCOM

This is my first attempt a creating a COM visible DLL. The intent
is to expose .NET Framework Regex methods in MS Office applications
via VBA.

Intended for use only with 64 bit MS Office. Build from Visual
Studio should automatically register the DLL and type library.
If it is not registering automatically, use the 64 bit regasm.exe
to register your created URL.

Note that Microsoft recently implement a new RegExp class in VBA. 
This new class beasically replicates to functionality of the 
VBScript library's RegExp class. For most users this will probably
be adequate, but for users looking for features that exist in the
.NET Framework implementation but not the VBScript implementation,
this DLL is an option.

Comments and suggestions are welcome, but please be kind to this
relative newbie.

To use with early binding, add a reference to the RegexCOM library 
from the VBA Editor:

<img width="446" height="359" alt="image" src="https://github.com/user-attachments/assets/223d5488-e182-4754-b0f1-f656bff09d9d" />

This will enable Intellisense, and also expose the members of the
RegexCOM namespace in the Object Browser. Filter for the RegexCOM
library, and you can see the RegexOptionsValue enum and the RegX
class. Select either to see the members of each:

<img width="491" height="286" alt="image" src="https://github.com/user-attachments/assets/22c2b40e-6452-4d31-8192-55dfc7038aa7" />

<img width="338" height="217" alt="image" src="https://github.com/user-attachments/assets/e0a4a03e-72be-4c37-b330-8e6a9df0b879" />

You can also use late binding, via CreateObject:

<img width="475" height="242" alt="image" src="https://github.com/user-attachments/assets/f9ada882-4543-467a-a265-4261d3c0c3c6" />
