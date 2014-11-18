# Begin of ProxyCommand for command: New-Object
Function New-Object {
<#
.SYNOPSIS
	Creates an instance of a Microsoft .NET Framework or COM object.
  Note:
    This is a proxy command to extend the origin cmdlet: Microsoft.PowerShell.Utility\New-Object
    For the origin Help do this: Get-Help Microsoft.PowerShell.Utility\New-Object
	
.DESCRIPTION
	The New-Object cmdlet creates an instance of a .NET Framework or COM object.

	You can specify either the type of a .NET Framework class or a ProgID of a COM object. By default, you type the fully qualified name of a .NET Framework class and the cmdlet returns a reference to an instance of that class. To create an instance of a COM object, use the ComObject parameter and specify the ProgID of the object as its value.

  Note:
    This is a proxy command to extend the origin cmdlet: Microsoft.PowerShell.Utility\New-Object
    For the origin Help do this: Get-Help Microsoft.PowerShell.Utility\New-Object
	
.PARAMETER ArgumentList
	Specifies a list of arguments to pass to the constructor of the .NET Framework class. Separate elements in the list by using commas (,). The alias for ArgumentList is Args.
	
.PARAMETER ComObject
	Specifies the programmatic identifier (ProgID) of an COM object to create,a class identifier (ClsID) (a string representation of a Guid) of an COM object to create or a filepath to an persistent COM object to create (like *.xlsx, *.doc or *.pdf).
  If you provide a ClsID or an filepath to create a COM Object the value of the -Property parameter is ignored and a warning is displayed.

  In case of a filepath the .NET class Microsoft.VisualBasic.GetObject() method is used to create the instance of the COM object.
  This method has other special capabilities to instantiate a COM object which are not used here.
  For more information see documentation of this method in the Microsoft Developer Network (MSDN).
  http://msdn.microsoft.com/en-us/library/e9waz863%28v=vs.90%29.aspx
	
.PARAMETER Property
	Sets property values and invokes methods of the new object.
	
	Enter a hash table in which the keys are the names of properties or methods and the values are property values or method arguments. New-Object creates the object and sets each property value and invokes each method in the order that they appear in the hash table.
	
	If the new object is derived from the PSObject class, and you specify a property that does not exist on the object, New-Object adds the specified property to the object as a NoteProperty. If the object is not a PSObject, the command generates a non-terminating error.
	
.PARAMETER Strict
	If the -ComObject parameter contains a filepath, a non-terminating error is generated when a the persistent COM object with that filepath is allready in the Running Object Table (ROT).
  In this case obtain the reference from the runnig object out of the ROT instead of creating a new object.
  
  	If the -ComObject parameter contains a Class Identifiers (ClsID) (the String representation of a Guid), a non-terminating error is generated when a the persistent COM object with that ClsID-Guid is allready in the Running Object Table (ROT).
  In this case obtain the reference from the runnig object out of the ROT instead of creating a new object.
  The ClsID-Guid can be match the Displayname or the ClassID attribute of the running object.
    
  Generates a non-terminating error when a COM object that you attempt to create uses an interop assembly. This feature distinguishes actual COM objects from .NET Framework objects with COM-callable wrappers.
	
.PARAMETER TypeName
	Specifies the fully qualified name of the .NET Framework class. You cannot specify both the TypeName parameter and the ComObject parameter.
	
.EXAMPLE
	PS C:\>New-Object -TypeName System.Version -ArgumentList "1.2.3.4"
	Major  Minor  Build  Revision
	
	-----  -----  -----  --------
	
	1      2      3      4
	This command creates a System.Version object. It uses a "1.2.3.4" string as the constructor.
	
.EXAMPLE
	PS C:\>$ie = New-Object -COMObject InternetExplorer.Application -Property @{Navigate2="www.microsoft.com"; Visible = $true}
	This command creates an instance of the COM object that represents the Internet Explorer application. The value of the Property parameter is a hash table that calls the Navigate2 method and sets the Visible property of the object to $true to make the application visible.
	This command is the equivalent of the following:
	$ie = New-Object -COMObject InternetExplorer.Application
	$ie.Navigate2("www.microsoft.com")
	$ie.Visible = $true
	
.EXAMPLE
	PS C:\>$a=New-Object -COMObject Word.Application -Strict -Property @{Visible=$true}
	New-Object : The object written to the pipeline is an instance of the type
	"Microsoft.Office.Interop.Word.ApplicationClass" from the component's primary
	interop assembly. If this type exposes different members than the IDispatch
	members, scripts written to work with this object might not work if the
	primary interop assembly is not installed.
	
	At line:1 char:14
	+ $a=New-Object  <<<< -COM Word.Application -Strict; $a.visible=$true
	This example demonstrates that adding the Strict parameter causes the New-Object cmdlet to generate a non-terminating error when the COM object uses an interop assembly.
	
.EXAMPLE
	The first command uses the ComObject parameter of the New-Object cmdlet to create a COM object with the "Shell.Application" ProgID. It stores the resulting object in the $objShell variable.
	PS C:\>$objshell = New-Object -COMObject "Shell.Application"
	
	The second command pipes the $objShell variable to the Get-Member cmdlet, which displays the properties and methods of the COM object. Among the methods is the ToggleDesktop method.
	PS C:\>$objshell | Get-Member
	TypeName: System.__ComObject#{866738b9-6cf2-4de8-8767-f794ebe74f4e}
	
	Name                 MemberType Definition
	
	----                 ---------- ----------
	
	AddToRecent          Method     void AddToRecent (Variant, string)
	
	BrowseForFolder      Method     Folder BrowseForFolder (int, string, int, Variant)
	
	CanStartStopService  Method     Variant CanStartStopService (string)
	
	CascadeWindows       Method     void CascadeWindows ()
	
	ControlPanelItem     Method     void ControlPanelItem (string)
	
	EjectPC              Method     void EjectPC ()
	
	Explore              Method     void Explore (Variant)
	
	ExplorerPolicy       Method     Variant ExplorerPolicy (string)
	
	FileRun              Method     void FileRun ()
	
	FindComputer         Method     void FindComputer ()
	
	FindFiles            Method     void FindFiles ()
	
	FindPrinter          Method     void FindPrinter (string, string, string)
	
	GetSetting           Method     bool GetSetting (int)
	
	GetSystemInformation Method     Variant GetSystemInformation (string)
	
	Help                 Method     void Help ()
	
	IsRestricted         Method     int IsRestricted (string, string)
	
	IsServiceRunning     Method     Variant IsServiceRunning (string)
	
	MinimizeAll          Method     void MinimizeAll ()
	
	NameSpace            Method     Folder NameSpace (Variant)
	
	Open                 Method     void Open (Variant)
	
	RefreshMenu          Method     void RefreshMenu ()
	
	ServiceStart         Method     Variant ServiceStart (string, Variant)
	
	ServiceStop          Method     Variant ServiceStop (string, Variant)
	
	SetTime              Method     void SetTime ()
	
	ShellExecute         Method     void ShellExecute (string, Variant, Variant, Variant, Variant)
	
	ShowBrowserBar       Method     Variant ShowBrowserBar (string, Variant)
	
	ShutdownWindows      Method     void ShutdownWindows ()
	
	Suspend              Method     void Suspend ()
	
	TileHorizontally     Method     void TileHorizontally ()
	
	TileVertically       Method     void TileVertically ()
	ToggleDesktop        Method     void ToggleDesktop ()
	
	TrayProperties       Method     void TrayProperties ()
	
	UndoMinimizeALL      Method     void UndoMinimizeALL ()
	
	Windows              Method     IDispatch Windows ()
	
	WindowsSecurity      Method     void WindowsSecurity ()
	
	WindowSwitcher       Method     void WindowSwitcher ()
	
	Application          Property   IDispatch Application () {get}
	
	Parent               Property   IDispatch Parent () {get}
	
	The third command calls the ToggleDesktop method of the object to minimize the open windows on your desktop.
	PS C:\>$objshell.ToggleDesktop()
	This example shows how to create and use a COM object to manage your Windows desktop.
	
.NOTES
	New-Object provides the most commonly-used functionality of the VBScript CreateObject function. A statement like Set objShell = CreateObject("Shell.Application") in VBScript can be translated to $objShell = New-Object -COMObject "Shell.Application" in Windows PowerShell.
  This version of New-Object provides the even the commonly-used functionality of the VBScript GetObject function. A statement like Set Set objExcelFile = GetObject("C:\Scripts\Test.xls") in VBScript can now be translated to $objExcelFile = New-Object -COMObject "C:\Scripts\Test.xls" in Windows PowerShell.
	
	New-Object expands upon the functionality available in the Windows Script Host environment by making it easy to work with .NET Framework objects from the command line and within scripts.

  Improved by Author: Peter Kriegel
  Current Version: 1.0.0 from: 17.November.2014
	
.INPUTS
	None
	
.OUTPUTS
	Object
	
.LINK
	http://go.microsoft.com/fwlink/p/?linkid=293993
	
.LINK
	Online Version:
	
.LINK
	Compare-Object
	
.LINK
	ForEach-Object
	
.LINK
	Group-Object
	
.LINK
	Measure-Object
	
.LINK
	Select-Object
	
.LINK
	Sort-Object
	
.LINK
	Tee-Object
	
.LINK
	Where-Object
#>


	[CmdletBinding(DefaultParameterSetName='Net')]
 	param(
 	    [Parameter(ParameterSetName='Net', Mandatory=$true, Position=0)]
 	    [string]
 	    $TypeName,
 	
 	    [Parameter(ParameterSetName='Com', Mandatory=$true, Position=0)]
 	    [string]
 	    $ComObject,
 	
 	    [Parameter(ParameterSetName='Net', Position=1)]
 	    [Alias('Args')]
 	    [System.Object[]]
 	    $ArgumentList,
 	
 	    [Parameter(ParameterSetName='Com')]
 	    [switch]
 	    $Strict,
 	
 	    [System.Collections.IDictionary]
 	    $Property)


 	end {

      # do COM object processing only if the parameter -ComObject has a value
      If(-Not [String]::IsNullOrEmpty($ComObject)) {

        # Try to create an COM object from a ClsID Guid
        Try{
          # Test if -ComObject parameter contains a GUID (ClsID)
          $Guid = $Null
          $Guid = [System.Guid]$ComObject
          If($Guid) {
            
            If($Property) {
              Write-Warning 'The -Property parameter is not supported in combination with the -ComObject parameter and ClsID-Guid!'
            }
            
              # If -Strict is present test if the object is allready loded into the Running Object Table (ROT)
              If($Strict.IsPresent) {
                # Test if the object ClsID-Guid is already loaded in Runnung Object Table (ROT)
                Get-ComRot | ForEach-Object {
                  $RunningObjectTableComponentInfo = $_

                  # Try to extract the a GUID out of the displayname
                  $RotDisplayGuid = ''
                  If($RunningObjectTableComponentInfo.DisplayName -match '(?:.*?)(\{{0,1}[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}{0,1})(?:.*)') {
                    # if the Dispaynam contains a GUID we use it to find the ProgID
                    Try {
                      $RotDisplayGuid = ([Guid]$Matches[1]).ToString()
                    } Catch { $RotDisplayGuid = ''}
                  }
                          
                  # if the ClsID-Guid matches the DisplayName or the ClassID of an object out of the Runnning Object Table (ROT) we show an error and return 
                  If(($Guid.ToString() -eq $RotDisplayGuid) -or ($Guid.ToString() -eq $RunningObjectTableComponentInfo.ClassID.ToString())) {
                    $message = "An object with the ClsId-Guid : $ComObject allready exist in the Running Object Table (ROT)!"
                    $exception = Microsoft.PowerShell.Utility\New-Object -TypeName System.IO.IOException -ArgumentList $message
                    $errorID = 'ComObjectAllreadyExist'
                    $errorCategory = [Management.Automation.ErrorCategory]::ResourceExists
                    $target = $ComObject
                    $errorRecord = Microsoft.PowerShell.Utility\New-Object Management.Automation.ErrorRecord $exception, $errorID,$errorCategory,$target
                    $PSCmdlet.WriteError($errorRecord)
                    Return
                  }
              }
            }
           
            Write-Verbose "Creating COM object from ClsID-Guid: $($Guid.ToString())"
            Try {
              # If conversion to a GUID was Succsesfull we try to create an COMObject from GUID
              Return [System.Activator]::CreateInstance(([System.Type]::GetTypeFromCLSID($Guid, $True)))
            } catch {
              $PSCmdlet.WriteError($_)
              Return
            }
          }
        } Catch {
          # -ComObject parameter seems not to be a Guid, dont throw error here 
          # $PSCmdlet.WriteError($_)
        }
                
        # test if -ComObject parameter contains a File an the file exist
        Try{
            $File = $Null
            # if PowerShell is currently not in the filesystem we must make shure that we work with the filesystem by use of the FileSystem provider
            If($ComObject -like '*FileSystem::*') {
                # using Get-Item to create a meaningful error message on error
                $File = Get-Item -Path $ComObject -ErrorAction Stop
            } Else {
                # Make shure that we work with the Filesystem by adding the FileSystem provider to the path
                If(Test-Path ('Microsoft.PowerShell.Core\FileSystem::' + $ComObject)) {
                  # using Get-Item to create a meaningful error message on error
                  $File = Get-Item -Path ('Microsoft.PowerShell.Core\FileSystem::' + $ComObject) -ErrorAction Stop
                }
                
            }
        } Catch {
          # if file is locked ore other Error throw error here
          $PSCmdlet.WriteError($_)
          return
        }
        # Try to create COM Object from file (a persistent COM Object)
        If($File) {
          
          If($Property) {
              Write-Warning 'The -Property parameter is not supported in combination with the -ComObject parameter and filepath!'
          }

          # If -Strict is present test if the object is allready loded into the Running Object Table (ROT)
          If($Strict.IsPresent) {
            # Test if the File is already loaded in Runnung Object Table (ROT), if so display an error and return
            If(Get-ComRot | Where-Object { $_.DisplayName -eq ($File.Fullname) }) {
              $message = "The file : $ComObject allready exist in the Running Object Table (ROT)!"
              $exception = Microsoft.PowerShell.Utility\New-Object -TypeName System.IO.IOException -ArgumentList $message
              $errorID = 'ComObjectAllreadyExist'
              $errorCategory = [Management.Automation.ErrorCategory]::ResourceExists
              $target = $ComObject
              $errorRecord = Microsoft.PowerShell.Utility\New-Object Management.Automation.ErrorRecord $exception, $errorID,$errorCategory,$target
              $PSCmdlet.WriteError($errorRecord)
              Return
            }
          }
                   
          Write-Verbose "Creating new COM object from File: $($File.Fullname)"
          Try {
            Add-Type -AssemblyName Microsoft.VisualBasic
            return [Microsoft.VisualBasic.Interaction]::GetObject(($File.Fullname))
          } Catch {
            $PSCmdlet.WriteError($_)
            return
          }
        }
      } 
      
      # if -ComObject parameter is not used or
      # -ComObject parameter in no Guid and no Filepath
      # we call the orign New-Object cmdlet to do normal Object creation
      Microsoft.PowerShell.Utility\New-Object @PSBoundParameters
   }

 
} # End function New-Object
