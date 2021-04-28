## Dynamic Start Menu Builder ##
### [Invoke-PowerShellBuilder.ps1](https://github.com/KryptykHermit/KryptykHermit.github.io/blob/main/PowerShell/StartMenuBuilderV2/Invoke-StartMenuBuilder.ps1) ###
This is my take on a dynamic Start Menu builder for Windows 10.  It will probably work on Windows 7 and 8 with a little adjustment.

The script walks through the following steps in creating the Start Menu
 1.   Specify the location of the new StartMenuLayout.xml file
 2.   Set the number of columns
 2.   Specify the group name(s)
 3.   Provide a priority driven regular expression grouping for each pin to be created (see note #1)
 4.   Acquire the current list of Start Menu applications
 5.   Create the sections of the XML and save to the location provided
 6.   Kill the StartMenuExperienceHost process to refresh the layout

>**NOTE #1**
>	-   The first regular expression is read in each array, and detects the applications named within the Get-StartApps PowerShell command.
>	- 	A pin is created using a 2x2 size starting at position 0,0
>	- 	The index position is incremented by 2
>	- 	The next application is read, and placed at 0,2
>	- 	The third application is read and positioned at 0,4
>	- 	The index position is incremented by 2, which exceeds the group limit space of 6 (0 to 5), so the column is reset to set, but the row is incremented by 2 
>	- 	The forth application is read, and placed at 2,0 and so on.

If an application in your regex array is not found, it is skipped and moves on to the next application.  If the application is installed at a later time, the pin is placed in the order you specify.  

```powershell
[string[]]$groupOrder2NoSub = '^Outlook(\s\d{4})?$', '^Word(\s\d{4})?$', '^Excel(\s\d{4})?$', '^PowerPoint(\s\d{4})?$', '^OneNote(\s\d{4})?$', '^Access(\s\d{4})?$', '^Publisher(\s\d{4})?$', '^Project(\s\d{4})?$', '^Visio(\s\d{4})?$'  
```

![Default Start Menu](https://kryptykhermit.github.io/PowerShell/StartMenuBuilderV2/StartMenuOverview.jpg)
