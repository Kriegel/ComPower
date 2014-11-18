Function Get-ComRotRegistered {
<#
.Synopsis
   Function to add inoformation out of the registry to a COM object out of the Runnig Object Table (ROT)

.DESCRIPTION
   
   Function to adds information out of the registry to a COM object out of the Runnig Object Table (ROT)
   This function is a helper function for the Get-ComRot function.
   This function adds 2 properties to ech processed Pamk_COM_ROT.RunningObjectTableComponentInfo object.
    RegisteredProgID
    ProgIDSources

   This function can add registry informations as a full object out of the Get-ComRegistered collection by use of the -AddObject parameter.
   You can find this objects as a list in the property 'RegisteredProgID' of each processed Pamk_COM_ROT.RunningObjectTableComponentInfo object.

   By default this function fills the 'RegisteredProgID' with a simple list of ProgID strings relate to the  a Pamk_COM_ROT.RunningObjectTableComponentInfo object.
   This function adds the following simple ProgID information out of the registry to a object returned by the Get-ComRot function by default:
  
    RegisteredProgID
      A String array of all Programmatic Identifier (ProgID) matching the ClsID of this object

    ProgIDSources
      String array with information which property of the object matches a ProgID in the RegisteredProgID array.
      This can have the following values:
        
        'ClsID(Guid)'
          The ClsID of the object matches a ClsID in the registry
          So the ProgID and the VersionIndependentProgID of this registry entry is added to the RegisteredProgID array
        
        'DisplayName(Guid)'
          If the DisplayName contains Guid the Guid is extracted out of the displayname and this guid matches a ClsID out of the registry.
          So the ProgID and the VersionIndependentProgID of this registry entry is added to the RegisteredProgID array

        'DisplayName(ProgID)'
          The DisplayName of the object contains ProgID and this ProgID matches a ProgID out of the registry.
          So the ProgID and the VersionIndependentProgID of this registry entry is added to the RegisteredProgID array

        'DisplayName(VersionIndependentProgID)'
          The DisplayName of the object contains ProgID and this ProgID matches a VersionIndependentProgID out of the registry.
          So the ProgID and the VersionIndependentProgID of this registry entry is added to the RegisteredProgID array
 
 See even paramter description of the parameter -EnrichWithRegistry of the Get-ComRot function.

.PARAMETER RunningObjectTableComponentInfo
  The Pamk_COM_ROT.RunningObjectTableComponentInfo object to add Registry informations to.

.PARAMETER DoNotUseCache
 	If you like use fresch registry informations out of the registry use the -DoNotUseCache parameter of the functions.
	If you use the -DoNotUseCache parameter, the cache is also freshed up.

  	The Registry entrys are changing only if a new COM Component ist registered or unregistered (or Sofware was installed).
	This happens not very often so this function take use of an array as a cache mechanism.
	The advantage is, that to read informations out of the cache is faster then reading it out of the registry every time this function is called.
	If you like use fresch registry informations out of the registry use the -DoNotUseCache parameter of the functions.
	If you use the -DoNotUseCache of any function, the cache is also freshed up.

.PARAMETER AddObject
  Use this parameter to add the registry informations as a full object to the Pamk_COM_ROT.RunningObjectTableComponentInfo
  If you do not use this parameter only the ProgID informations are added to the Pamk_COM_ROT.RunningObjectTableComponentInfo object

.EXAMPLE
   Get-ComRot | Get-ComRotRegistered

.EXAMPLE
   Get-ComRot | Get-ComRotRegistered -AddObject 

.EXAMPLE
  Get-ComRotRegistered -RunningObjectTableComponentInfo (Get-ComRot | Where-Object { $_.DisplaName -eq 'Document1' })
  
.EXAMPLE
  Get-ComRotRegistered -RunningObjectTableComponentInfo (Get-ComRot | Where-Object { $_.DisplaName -eq 'Document1' }) -DoNotUseCache   

.Notes
  
  Dependencies:
  This Function uses the 'Get-ComRegistered' Function to find the registered COM components from Registry
  
  Author: Peter Kriegel
  Version 1.0.0. 7.November.2014
#>

  [CmdletBinding()]
  Param(
    [Parameter(Mandatory=$True, ValueFromPipeline=$True)]
    [Pamk_COM_ROT.RunningObjectTableComponentInfo[]]$RunningObjectTableComponentInfo,
    [Switch]$AddObject,
    [Switch]$DoNotUseCache
  )
  
  begin {
     If(-not $Script:ComRegisteredCache) {
      # The Registry entrys are changing only if a new COM Component ist registered or unregistered
      # This happens not very often so we use a Variable to hold the Registered COM components from Registry
      # as Array which is used as a static cache
      $Script:ComRegisteredCache = [System.Collections.ArrayList]@()
    }

  }  # end of begin block
  
  Process {
    
    ForEach($ROTComponentInfo in $RunningObjectTableComponentInfo) {
    $Guid = $Null
    [System.Collections.ArrayList]$ResultList = @()
    [System.Collections.ArrayList]$ProgIDMatches = @() 

        # Try to extract the a GUID out of the displayname
        If($ROTComponentInfo.DisplayName -match '(?:.*?)(\{{0,1}[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}{0,1})(?:.*)') {
          # if the Dispaynam contains a GUID we use it to find the ProgID
          $GUID = [Guid]$Matches[1]
        } else {
          # no GUID in Displayname found
          $GUID = $Null
        }

    If((-not $Script:ComRegisteredCache.Count -gt 0) -or $DoNotUseCache.IsPresent ) {
      # read the Registered COM components fresh from Registry into the Array which is used as a static cache
      [void]([System.Collections.ArrayList]$Script:ComRegisteredCache).AddRange((Get-ComRegistered -DoNotUseCache:($DoNotUseCache.IsPresent)))
    }

    ForEach($RegisteredCom in $Script:ComRegisteredCache) {
    
      $TextGuid = ([Guid]$RegisteredCom.ClsID).ToString()
      $RotClsIdTakenFrom = @()
      $ProgIDFound = $False

      If($Guid) {
        If($TextGuid -eq ($Guid.ToString())) {
          $RotClsIdTakenFrom += 'DisplayName(Guid)'
          If(-not $ProgIDMatches.Contains('DisplayName(Guid)')) {
            [void]$ProgIDMatches.Add('DisplayName(Guid)')
          }
          $ProgIDFound = $True
        }
      }

      If($TextGuid -eq (([Guid]$ROTComponentInfo.ClsID).ToString())) {
        $RotClsIdTakenFrom += 'ClsID(Guid)'
        If(-not $ProgIDMatches.Contains('ClsID(Guid)')) {
          [void]$ProgIDMatches.Add('ClsID(Guid)')
        }      
        $ProgIDFound = $True
      }

      If(((-not [String]::IsNullOrEmpty($RegisteredCom.ProgID)) -and ($ROTComponentInfo.DisplayName -Like "*$($RegisteredCom.ProgID)*")))
      {
        If(-not ($RotClsIdTakenFrom -contains 'DisplayName(ProgID)')) {
          $RotClsIdTakenFrom += 'DisplayName(ProgID)'
        }
        $RotClsIdTakenFrom += 
        If(-not $ProgIDMatches.Contains('DisplayName(ProgID)')) {
          [void]$ProgIDMatches.Add('DisplayName(ProgID)')
        }
        $ProgIDFound = $True
      }

      If((-not [String]::IsNullOrEmpty($RegisteredCom.VersionIndependentProgID)) -and ($ROTComponentInfo.DisplayName -Like "*$($RegisteredCom.VersionIndependentProgID)*"))
      {
        If(-not ($RotClsIdTakenFrom -contains 'DisplayName(VersionIndependentProgID)')) {
          $RotClsIdTakenFrom += 'DisplayName(VersionIndependentProgID)'
        }
        If(-not $ProgIDMatches.Contains('DisplayName(VersionIndependentProgID)')) {
          [void]$ProgIDMatches.Add('DisplayName(VersionIndependentProgID)')
        }
      }
        
      If($ProgIDFound) {
        If($AddObject.IsPresent) {

          If(-not [String]::IsNullOrEmpty($RotClsIdTakenFrom)) {
            [void]$ResultList.Add(($RegisteredCom | Select-Object * ,@{Name='RotClsIdTakenFrom';Expression={$RotClsIdTakenFrom}}))
          }
        } else {
          If(-not $ResultList.Contains($RegisteredCom.ProgID)-and (-not [String]::IsNullOrEmpty($RegisteredCom.ProgID))) {
            [void]$ResultList.Add($RegisteredCom.ProgID)
          }
          If(-not $ResultList.Contains($RegisteredCom.VersionIndependentProgID) -and (-not [String]::IsNullOrEmpty($RegisteredCom.VersionIndependentProgID))) {
            [void]$ResultList.Add($RegisteredCom.VersionIndependentProgID)
          }
        }

      }

    }

    # add Properties to the inputobject withthe help of Select-Object because it is faster then Add-Member
    $ROTComponentInfo | Select-Object * ,@{Name='RegisteredProgID';Expression={$ResultList.ToArray()}},@{Name='ProgIDSources';Expression={$ProgIDMatches}}
    }
  } # end of process block

}