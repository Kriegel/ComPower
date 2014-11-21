ComPower
========

![alt text](https://github.com/Kriegel/ComPower/blob/master/ComPower.png "ComPower Logo")

##TOPIC

ComPower Module Overview

Author: Peter Kriegel
Version 1.0.0. 13.November.2014

##SHORT DESCRIPTION

ComPower is a Windows PowerShell module to work with the Component Object Model (COM).
COM is also known under the terms “Object Linking and Embedding” (OLE) and ActiveX.

##LONG DESCRIPTION

For additional informations read : Get-Help 'about_ComPower' ; after loading the module!

ComPower is a Windows PowerShell module to work with the Component Object Model (COM).
COM is also known under the terms “Object Linking and Embedding” (OLE) and ActiveX.

This Module works manly with COM classes (called  components here) which are registerd inside the Windows Registry,
or running objects (moniker) wich can be found in the Running Object Table (ROT) of a system.

In this version of the ComPower module, Registration-free COM components and Interfaces are not supported.
Adam Driscoll has written a Blog post about: Using Unregistered COM Objects in PowerShell

see: http://csharpening.net/?p=1427

The registry tracks where components are deployed on a local System and remote hosts.
COM classes, interfaces and type libraries are listed by GUIDs called Class Identifiers (ClsID) in the registry,
under the registry path HKEY_CLASSES_ROOT\CLSID or HKEY_CLASSES_ROOT\Wow6432Node\CLSID for classes and HKEY_CLASSES_ROOT\Interface or HKEY_CLASSES_ROOT\Wow6432Node\interface for interfaces.
Interfaces are not considered within this module.

The registry keys directly under HKEY_CLASSES_ROOT are not considered within this module.
COM libraries use the registry to locate either the correct local libraries for each COM object or the network location for a remote service.
The Regsvr32.exe (Microsoft Register Server) command-line tool is a command-line utility in Microsoft Windows operating systems for registering and unregistering DLLs and ActiveX controls out of .dll, .ocx or .exe files in the Windows Registry.	

####Cache the registry entrys on load of the module

The Registry entrys are changing only if a new COM Component ist registered or unregistered (or Sofware was installed).
This happens not very often so this ComPower module take use of an array as a cache mechanism.
This array is filled up on load of the ComPower module, so loading the module takes a while.
Nearly all functions out of this module are working with this information out of this cache.
The advantage is, to read informations out of the cache is faster then reading out of the registry.
If you like to use fresch registry informations out of the registry use the -DoNotUseCache parameter of the functions.
If you use the -DoNotUseCache of any function, the cache is also freshed up.

---------------------------------------------------------------------------

##This module contains the following function (script-cmdlets):

###Get-ComRegistered

Function to list all registered COM classes of a system from the registry key CLSID.
This Function is a simple way to replace tools like: OLE/COM Object Viewer (oleview.exe), RegDllView, OLE/COM Object Explorer, Objektbrowser or ActiveXHelper

###Get-ComRot

Function to list all COM objects out of the Running Object Table (ROT) and get direct access to the Running Object Table (ROT)
With the help of this function you can also get a reference of an COM object out of the ROT elective.
The Running Object Table (ROT) is a machine-wide table in which objects can register themselves.

###Get-ComRotIntance

Function to obtain a reference to the instance of a COM object out of the Running Object Table (ROT)
So this method is similar to the Visual Basic method GetObject().
The advantage ot this function over the Visual Basic method GetObject() is the freedom to choose which object you like to have (direct access to the object reference). 

###Get-ComRotRegistered

Function to add inoformation out of the registry to a COM object out of the Runnig Object Table (ROT)
This function is a helper function for the Get-ComRot function.
This function adds ProgID information out of the registry to a object returned by the Get-ComRot function

###New-Object

Proxy-function to extend the origin New-Object Cmdlet.
The origin New-Object Cmdlet can instantiate COM objects only over their Programmatic Identifier (ProgID).
Some COM objects do not have a ProgID or the ProgID is ambiguous so you have problems to get this COM object.
With this New-Object function you can now create an instance of a COM object from ClsID or a file path.

###Remove-ComObject

Function to help to releases a COM object from the Windows PowerShell process space.
Because the COM objects live in unmanage memory and Windows PowerShell lives in managed .NET memory the binding to each other is VERY fragile!
So it is very hard to clean up a COM object from .NET (Windows PowerShell).
This function trys to help you to get rid of a COM object.

>Bevor you create a COM object with .NET or Windows PowerShell!
>Take my advice! Allways think twice!
>&nbsp;&nbsp;&nbsp;&nbsp;Spirits that I've cited<br />
>&nbsp;&nbsp;&nbsp;&nbsp;My commands ignore.<br />
>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(Johann Wolfgang von Goethe)

---------------------------------------------------------------------------

##USAGE

The first thing is to get knowledge which Component Object Model COM objects are available on a Windows System.

###Get all COM classes available on a system

The function Get-ComRegistered is to get all COM objects, that are registered in the registry keys:
"HKEY_CLASSES_ROOT\CLSID" and "HKEY_CLASSES_ROOT\Wow6432Node\CLSID"

Example call:

```powershell
# Get all registered COM objects out of the registry
Get-ComRegistered

# Get only registered COM objects which have a Programmatic Identifier ProgID 
Get-ComRegistered | Where-Object {$_.ProgID}

# Get registered COM objects which have the Programmatic Identifier like Excel.Application 
Get-ComRegistered | Where-Object {$_.ProgID -like '*Excel.Application*'}

# Get registered COM objects which have the Programmatic Identifier like *Word* or the friendly classname *Word* 
Get-ComRegistered | Where-Object {$_.ProgID -like '*Word*' -or $_.FriendlyClassName -like '*Word*' }

# Get registered COM objects which have the Class Identifiers (ClsID) guid like ‘*000209ff-0000-0000-c000-000000000046*’
Get-ComRegistered | Where-Object {$_.ClsID -like '*000209ff-0000-0000-c000-000000000046*'}
```
---

The second thing you probably like to do is to create an instance of an Component Object Model COM object and get a reference to it with Windows PowerShell.

###Create new COM objects

In Visual Basic the method CreateObject() and GetObject() are used for that.

This module uses a proxy command to extend the origin cmdlet: Microsoft.PowerShell.Utility\New-Object to create instances of COM object 3 different ways:

1). Create an instance from a persistent COM object from disk by use of a file path.
(Microsoft Office documents are persistent COM objects on disk *.xlsx, *.docx, *.pptx …)
```powershell
# Create an Excel Workbook object from file path, use –Strict to check if document is already running
$XlWorkb = New-Object -ComObject 'C:\users\Jon\Doc\Chart.xlsx' -Verbose –Strict
```
2). Create an instance of a COM object by use of its Class Identifier ClsID.
(ClsID is a worldwide unique GUID which identifies a COM object)
```powershell
# Create an Microsoft Word Application object from ClsID-Guid use –Strict to check if Word is already running
$WordApp = New-Object -ComObject '000209ff-0000-0000-c000-000000000046' -Verbose –Strict
```
3). Create an instance of a COM object by use of its Programmatic Identifier (ProgID).
(A ProgID is one or more alias name(s) to an COM object. Like the CLSID, the ProgID identifies a class, but with less precision.)
```powershell
# Create an PowerPoint Applicatin object from ProgID, use –Strict to check if interop assembly is used
$XlWorkb = New-Object -ComObject 'PowerPoint.Application' -Verbose -Strict
```
(For more information see in the documentation of the associated functions of this ComPower module)

---

###Get a list ot the running objects out of the Running Object Table (ROT)

If you like to know which COM objects currently running on y a system, take a look into Running Object Table (ROT).

For this you can use the Get-ComRot function:
```powershell
# Get information of all running objects out of the Running Object Table (ROT)
Get-ComRot

# Get information of all running objects out of the Running Object Table (ROT) with ProgID informations out of the registry
Get-ComRot -EnrichWithRegistry
```

The Get-ComRot function returns a List of Pamk_COM_ROT.RunningObjectTableComponentInfo objects.
The Pamk_COM_ROT.RunningObjectTableComponentInfo is a custom .NET class and
it is similar to the System.Runtime.InteropServices.ComTypes.IMoniker interface but the Pamk_COM_ROT.RunningObjectTableComponentInfo is not bound to the COM object and has no reference to it.

---

### Get a reference to an running object out of the Running Object Table (ROT)

If you need a reference to an running object out of the Running Object Table (ROT) you can obtain it in 2 ways:
(This method is similar to the Visual Basic method GetObject() but you have direct acces to each object.)

1). use the GetInstance() method of a Pamk_COM_ROT.RunningObjectTableComponentInfo object returned by the Get-ComRot function
```powershell
# Get a moniker (a reference) to an Microsoft word document COM object with the DisplayName 'Document1' out of the Running Object Table (ROT)
$WordDoc = (Get-ComRot | Where-Object { $_.DisplayName -eq 'Document1'}).GetInstance()
```
2). You can use the Get-ComRotInstance function to obtain a reference
```powershell
# Get a moniker (a reference) to an Microsoft word document COM object with the DisplayName 'Document1' out of the Running Object Table (ROT)
$WordDoc = Get-ComRot | Where-Object { $_.DisplayName -eq 'Document1'} | Get-ComRotIntance
```

###Remove a COM object from the Windows PowerShell process

You can clean up a COM object (moniker) with the help of the Remove-ComObject
Because the COM objects live in unmanage memory and Windows PowerShell lives in managed .NET memory the binding to each other is VERY fragile!
So it is very hard to clean up a COM object from .NET (Windows PowerShell).
Please read the full help of this function!

>Bevor you create a COM object with .NET or Windows PowerShell!
>Take my advice! Allways think twice!
>&nbsp;&nbsp;&nbsp;&nbsp;Spirits that I've cited<br />
>&nbsp;&nbsp;&nbsp;&nbsp;My commands ignore.<br />
>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(Johann Wolfgang von Goethe)
