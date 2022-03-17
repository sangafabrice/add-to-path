# **The Add-To-Path utility**
---
##### Author: Fabrice Sanga
<br/>
<br/>

Simple utility to modify the PATH environment variable on Windows. It makes use of context menus to access the variable and change it. The shell objects involved are the directory and the directory background which keys are located under *HKEY_CURRENT_USER\SOFTWARE\Classes\Directory* within the Windows registry.

The concept is simple, the environment variables are all located in *HKEY_CURRENT_USER\Environment\Path* when only available to the current user, and *HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\Path* for system-wide variables. So the script reads and updates the `Path` value names at the user's wish. The script handles the changes to the context menus accordingly.

The utility is built in a way that makes the updates of system variables silent. This means the User Access Control (UAC) does not prompt for administrator privileges. It proceeds uninterrupted.
<br/>
<br/>

This is how it looks:

__1. A folder that is not on the PATH variable__

![](https://drive.google.com/uc?export=view&id=1vJ5F88dQM_Zt8SPtUKytNZaP2AVBr74a)
<br/>
<br/>

__2. A folder that is on the PATH variable__

![](https://drive.google.com/uc?export=view&id=1TEI2vkHcZ9-vXizElNeC5P3ktCHZ_NJJ)

The conditional display of the options is made using Advanced Query Syntax (AQS) on the implemented verbs. It helps include and exclude directories. Another project will work on context handlers. The canonical queries make extensive use of the `System.ItemPathDisplay` property that connects to the folder path. Just what is needed. And a chain of conjunction and disjunction associations does the trick. It is a complicated way to mean `AND` and `OR` operators.
<br/>
<br/>

__3. The folder background__

![](https://drive.google.com/uc?export=view&id=1QdiBOf2pOIkYXc2XoAFSFv9DPAdTA2n7)
<br/>
<br/>

The utility allows to add a set of directories at the same time and the set is identified by a `PathID` (7z, git, ...). The checkmark notifies that all the folders in the set are on the `PATH`. *Reset all* reverts the environment variable to a predefined set of folders.
<br/>
<br/>

---

The usage of the script is as follows using a *Windows Scripting Host*, either `cscript` or `wscript` since it is all *VBScript*.

* To set up the utility.
```
wscript.exe set-path.vbs /install
```
* To register a set of directories
```
wscript.exe set-path.vbs PathID /install [/user:Path[;...]] /system:Path[;...]
wscript.exe set-path.vbs PathID /install /user:Path[;...] [/system:Path[;...]]

PathID      The set of paths identifiers
/user       Registers the set of directories as only 
            available to the current user
/system     Registers the set of directories as 
            available to every user
```
* To add or remove a group of directories with a path identifier
```
wscript.exe set-path.vbs [-|+]PathID [/elevate]

+           Adds the set of directories to PATH
            the option is the default.
-           Removes the set of directories from PATH
/elevate    to run as administrator
```
* To add or remove a group of directories with no path identifier
```
wscript.exe set-path.vbs -[FolderPath[;...]] [/path] [/elevate]
wscript.exe set-path.vbs [+][FolderPath[;...]] [/path[:{system|user}]] [/elevate]

FolderPath  The directory path to add or remove
```
* To reset path
```
wscript.exe set-path.vbs /reset [/elevate]
```