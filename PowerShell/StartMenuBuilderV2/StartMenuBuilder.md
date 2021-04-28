## Dynamic Start Menu Builder ##
### Invoke-PowerShellBuilder.ps1 ###
Here is my take on a dynamic Start Menu builder for Windows 10.  It will probably work on Windows 7 and 8 with a little adjustment.

The primary mechanism for building the Start Menu with this script is the Start Menu itself.  Most, if not all, programs will be installed in the **C:\ProgramData\Microsoft\Windows\Start Menu\Programs** folder, therefore to create the pinned applications, we utilize this folder and its contents as a source of "truth".



The script utilizes regular expression (regex) to fine tune the applications to be queried, and in some cases where multiple apps are captured, a way to provide the actual application name in addition.

Each set of regex set is priority based.  Application 1