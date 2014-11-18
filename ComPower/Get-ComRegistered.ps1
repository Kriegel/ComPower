Function Get-ComRegistered {                       
<#
.Synopsis            
   Gets all ClsID, ProgID and VersionIndependentProgID of Component Object Model (COM) Object registered on the system.
               
.Description            
   Gets all ClsID, ProgID and VersionIndependentProgID of Component Object Model (COM) Object registered on the system.
   The ProgIDs returned can be used with New-Object -comObject
   Because there are so many entrys in the Registry this function is optimized for speed not for beauty!

.PARAMETER ComputerName
  Specifies the target computer for the operation.
  Enter a fully qualified domain name, a NetBIOS name, or an IP address.
  When the remote computer is in a different domain than the local computer, the fully qualified domain name is required.

  The default is the local computer.
  To specify the local computer, such as in a list of computer names, use "localhost", the local computer name, or a dot (.).

.PARAMETER DoNotUseCache
 	If you like use fresch registry informations out of the registry use the -DoNotUseCache parameter of the functions.
	If you use the -DoNotUseCache parameter, the cache is also freshed up.

  	The Registry entrys are changing only if a new COM Component ist registered or unregistered (or Sofware was installed).
	This happens not very often so this function take use of an array as a cache mechanism.
	The advantage is, that to read informations out of the cache is faster then reading it out of the registry every time this function is called.
	If you like use fresch registry informations out of the registry use the -DoNotUseCache parameter of the functions.
	If you use the -DoNotUseCache of any function, the cache is also freshed up.

.Example            
   Get-ComRegistered

   Get all registered COM Object wich have a ProgID
   If a COM Object has no ProgID entry in the Registry it is skipped

.Example            
   Get-ComRegistered -All

   Get all registered COM Object out of the registry of the local System even if they have no ProgID

.Example            
   Get-ComRegistered -ComputerName 'Server1'

   Get all registered COM Object wich have a ProgID, out of the registry of the System with Name 'Server1'
   If a COM Object has no ProgID entry in the Registry it is skipped 

.OUTPUTS

  PScustom object with the following properties

  ProgID                   : the Programmatic Identifier (ProgID) of the registry entry
  ClsID                    : Class Identifiers (ClsID)  of the registry entry (the registry Key name)
  VersionIndependentProgID : Version independent Programmatic Identifier (ProgID) of the registry entry
  ComputerName             : source Computername of the registry entry
  ParentRegistryKey        : PSCustomObject with informations about the parent registry key
                            ChildName   : Name of the child registry key
                            IsContainer : allways True           
                            Name        : Name of the registry key
                            ParentPath  : Name of the parent registry key
                            Path        : Path to the registry key
                            SubKeyCount : count of subkeys found in the registry
                            View        : allways Default 

  One or no of the following registry properties are pointing to the source .dll .ocx of the COM Object  

  InprocHandler            : 
  InprocHandler32          : 
  InprocServer32           : 
  LocalServer              : 
  LocalServer32            : 
  LocalService             :

.Notes
  Author: Peter Kriegel
  Version 1.0.0. 4.November.2014
  Credits:
  This one was inspired by a blog post by James Brundage:
  See : http://blogs.msdn.com/b/powershell/archive/2009/03/20/get-progid.aspx 
#> 

    [CmdletBinding()]
    param(
      [String]$ComputerName = $Env:ComputerName,
      [Switch]$DoNotUseCache
    )  

  begin {

   If(-not $Script:ComRegisteredCache) {
      # The Registry entrys are changing only if a new COM Component ist registered or unregistered
      # This happens not very often so we use a Variable to hold the Registered COM components from Registry
      # as Array which is used as a static cache
      $Script:ComRegisteredCache = [System.Collections.ArrayList]@()
    }
    Function InternalWorkHorse {
      
      [CmdletBinding()]
      param(
        [String]$ComputerName
      )

      Try {
	      $HiveRegistryKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot,$ComputerName,[Microsoft.Win32.RegistryView]::Default)
      } Catch {
	      Write-Error $_
	      Return $Null
      }
      
      # Create array with names of registry keys to visit
      $ClsidRegistryKeys = @('CLSID')

      # test if the Wow6432Node key exist
      Try{
        $SubKey = $HiveRegistryKey.OpenSubKey('Wow6432Node\CLSID123')
        $SubKey.GetValue('')
        # Wow6432Node exist add to list
        $ClsidRegistryKeys += 'Wow6432Node\CLSID'
      } Catch {}


      # clear the cache to refill it with this function
      ([System.Collections.ArrayList]$Script:ComRegisteredCache).Clear()

      ForEach($ClsidRegistryKey in $ClsidRegistryKeys) {

            $ClsidRegistryKeyChildName = ($ClsidRegistryKey -split '\\')[(($ClsidRegistryKey -split '\\').count -1)]

            $SubKeyNames = $HiveRegistryKey.OpenSubKey($ClsidRegistryKey).GetSubKeyNames()

            $ParentRegistryKey = New-Object -TypeName PsObject -Property @{
                          Path = "HKEY_CLASSES_ROOT\$ClsidRegistryKey"
                          ParentPath = 'HKEY_CLASSES_ROOT'
                          ChildName = $ClsidRegistryKeyChildName
                          IsContainer = $True
                          SubKeyCount = $SubKeyNames.Count
                          View = 'Default'
                          Name = "HKEY_CLASSES_ROOT\$ClsidRegistryKey"
                        }
            
            Write-Verbose "Processing Registry Key: HKEY_CLASSES_ROOT\$ClsidRegistryKey" 
            
            # process each key inside the Registry
            ForEach($SubKeyName in $SubKeyNames)
            {

              Write-Verbose "Processing Sub-Registry Key: HKEY_CLASSES_ROOT\$ClsidRegistryKey\$SubKeyName" 

              # create a new empty Object to return the result
              # because New-Object is very slow to create a new PSObject we use the '' | Select-Object trick here
              # '' | Select-Object trick here; is a fast way to create a new empty Object with PowerShell 2.0
              $ResultObject = '' | Select-Object ProgID,ClsID,VersionIndependentProgID,ComputerName,ParentRegistryKey,FriendlyClassName,InprocHandler,InprocHandler32,InprocServer32,LocalServer,LocalServer32,LocalService
              $ResultObject.PStypenames.Clear()
              $ResultObject.PStypenames.Add('System.Management.Automation.PSCustomObject')
              $ResultObject.PStypenames.Add('System.Object')
    
              Try{
                $SubKey =  $HiveRegistryKey.OpenSubKey("$ClsidRegistryKey\$SubKeyName\ProgId")
                $ProgId = $SubKey.GetValue('')
              }
              Catch { $ProgId = '' }

              $ResultObject.ProgID = $ProgId
              $ResultObject.ParentRegistryKey = $ParentRegistryKey
              $ResultObject.Computername = $ComputerName

              Try{
                $SubKey =  $HiveRegistryKey.OpenSubKey("$ClsidRegistryKey\$SubKeyName\VersionIndependentProgID")
                $ResultObject.VersionIndependentProgID = $SubKey.GetValue('')
              }
              Catch { $ResultObject.VersionIndependentProgID = '' }

              Try{
                $ResultObject.ClsID = "{$(([Guid]$SubKeyName).ToString())}"
              }
              Catch {
                Write-Warning "ClsID is Empty on ProgID: $ProgId RegistryKey: HKEY_CLASSES_ROOT\$ClsidRegistryKey\$SubKeyName " 
                continue
              }

              Try{
                $SubKey =  $HiveRegistryKey.OpenSubKey("$ClsidRegistryKey\$SubKeyName")
                $ResultObject.FriendlyClassName = $SubKey.GetValue('')
              }
              Catch { $ResultObject.FriendlyClassName = '' }

              Try{
                $SubKey =  $HiveRegistryKey.OpenSubKey("$ClsidRegistryKey\$SubKeyName\InprocHandler")
                $ResultObject.InprocHandler = $SubKey.GetValue('')
              }
              Catch { $ResultObject.InprocHandler = '' }

              Try{
                $SubKey =  $HiveRegistryKey.OpenSubKey("$ClsidRegistryKey\$SubKeyName\InprocHandler32")
                $ResultObject.InprocHandler32 = $SubKey.GetValue('')
              }
              Catch { $ResultObject.InprocHandler32 = '' }

              Try{
                $SubKey =  $HiveRegistryKey.OpenSubKey("$ClsidRegistryKey\$SubKeyName\InprocServer32")
                $ResultObject.InprocServer32 = $SubKey.GetValue('')
              }
              Catch { $ResultObject.InprocServer32 = '' }

              Try{
                $SubKey =  $HiveRegistryKey.OpenSubKey("$ClsidRegistryKey\$SubKeyName\LocalServer")
                $ResultObject.LocalServer = $SubKey.GetValue('')
              }
              Catch { $ResultObject.LocalServer = '' }

              Try{
                $SubKey =  $HiveRegistryKey.OpenSubKey("$ClsidRegistryKey\$SubKeyName\LocalServer32")
                $ResultObject.LocalServer32 = $SubKey.GetValue('')
              }
              Catch { $ResultObject.LocalServer32 = '' }

              Try{
                $SubKey =  $HiveRegistryKey.OpenSubKey("$ClsidRegistryKey\$SubKeyName\LocalService")
                $ResultObject.LocalService = $SubKey.GetValue('')
              }
              Catch { $ResultObject.LocalService = '' }

            # return resulting Object and 
            # read the Registered COM components fresh from Registry into the Array which is used as a static cache
            [void]([System.Collections.ArrayList]$Script:ComRegisteredCache).Add($ResultObject)
            $ResultObject

            } # end of ForEach($SubKeyName in $SubKeyNames) 32Bit hive

      }
    
    } # end of function InternalWorkHorse block
  
  } # end of begin block
  process {
    # nothing here!
  } # end of process block
  end {

    # (re)read values from registry if the cache is empty or the cache should not be used or the All setting has changed
    If((-not $Script:ComRegisteredCache.Count -gt 0) -or $DoNotUseCache.IsPresent ) {
      
      Write-Verbose 'Reading the Registry entrys to the cache'
      # read the Registered COM components fresh from Registry 
      InternalWorkHorse -ComputerName $ComputerName

    } else {
    
      Write-Verbose 'Using the cache to output the Result'
      # return Result from cache
      $Script:ComRegisteredCache
    }
    
  } # end of end block
}