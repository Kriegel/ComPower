$TypeDefinition = @'
using System;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

/*
 * Credits to the orginators:
 * Orginal Code from Rainbird Year: 2007
 * See: http://dotnet-snippets.de/snippet/laufende-com-objekte-abfragen/526
 * Modified by: Sebastian.Lange Year 2012
 * See: http://www.mycsharp.de/wbb2/thread.php?threadid=36340
 */

namespace Pamk_COM_ROT
{
    /// <summary>
    /// This class provides direct access to the Running Object Table (ROT) which holds the current runnig COM Objects
    /// The Running Object Table (ROT) is a machine-wide table in which objects can register themselves.
    /// Because it is optional for a COM object to register within the ROT you can not find all running COM Objects inside this table
    /// Even sometimes you will find only one entry in the ROT for an COM object which runs many instances. This is by design of the COM producer.  
    /// </summary>
    public static class RunningObjectTable
    {
        /// <summary>
        /// Private default constructor
        /// </summary>
        //private RunningObjectTable() { }

        // Win32 API call to read the ROT
        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable pprot);

        // Win32 API call to create the binding
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx pctx);

        public static List<RunningObjectTableComponentInfo> GetComponentsFromROT()
        {
            IEnumMoniker monikerList = null;
            IRunningObjectTable runningObjectTable = null;
            List<RunningObjectTableComponentInfo> resultList = new List<RunningObjectTableComponentInfo>();
            try
            {
                // query table and returns null if no objects runnings
                if (GetRunningObjectTable(0, out runningObjectTable) != 0 || runningObjectTable == null)
                    return null;

                // query moniker & reset
                runningObjectTable.EnumRunning(out monikerList);
                monikerList.Reset();

                IMoniker[] monikerContainer = new IMoniker[1];
                IntPtr pointerFetchedMonikers = IntPtr.Zero;

                // fetch all moniker
                while (monikerList.Next(1, monikerContainer, pointerFetchedMonikers) == 0)
                {
                    // create binding object
                    IBindCtx bindInfo;
                    CreateBindCtx(0, out bindInfo);

                    // query com proxy info      
                    object comInstance = null;
                    runningObjectTable.GetObject(monikerContainer[0], out comInstance);

                    string ppszDisplayName;
                    try { monikerContainer[0].GetDisplayName(bindInfo, null, out ppszDisplayName); }
                    catch { ppszDisplayName = ""; }

                    Guid pClassID;
                    try { monikerContainer[0].GetClassID(out pClassID); }
                    catch { pClassID = Guid.Empty; }
                    
                    System.Runtime.InteropServices.ComTypes.FILETIME pLastChangedFileTime;
                    try { runningObjectTable.GetTimeOfLastChange(monikerContainer[0], out pLastChangedFileTime);}
                    catch { pLastChangedFileTime = new System.Runtime.InteropServices.ComTypes.FILETIME(); }
                    
                    long pcbSize;
                    try { monikerContainer[0].GetSizeMax(out pcbSize);}
                    catch { pcbSize = 0;}
                    
                    
                    Boolean IsDirty;
                    try {IsDirty = (monikerContainer[0].IsDirty() == 0); }
                    catch {IsDirty = false;}

                    Boolean IsRunning;
                    try {IsRunning = (runningObjectTable.IsRunning(monikerContainer[0]) == 0); }
                    catch {IsRunning = false;}

                    // creating the unbound object to hold the Data out of the current component IMoniker
                    RunningObjectTableComponentInfo ROTComponent = new RunningObjectTableComponentInfo(
                        ppszDisplayName,
                        pClassID,
                        ConvertFromFILETIME(pLastChangedFileTime),
                        pcbSize,
                        TypeDescriptor.GetComponentName(comInstance, false),
                        TypeDescriptor.GetClassName(comInstance),
                        IsRunning,
                        IsDirty
                    );

                    resultList.Add(ROTComponent);

                    // clean up and release object 
                    Marshal.ReleaseComObject(comInstance);
                    Marshal.ReleaseComObject(bindInfo);
                }

                // not running
                return resultList;
            }
            finally
            {
                // release proxies
                if (runningObjectTable != null)
                    Marshal.ReleaseComObject(runningObjectTable);
                if (monikerList != null)
                    Marshal.ReleaseComObject(monikerList);
            }
        }

        public static object GetInstanceFromROT(RunningObjectTableComponentInfo RunningComponent)
        {
            IEnumMoniker monikerList = null;
            IRunningObjectTable runningObjectTable = null;
            object ResultInstance = null;

            try
            {
                // query table and returns null if no objects runnings
                if (GetRunningObjectTable(0, out runningObjectTable) != 0 || runningObjectTable == null)
                    return null;

                // query moniker & reset
                runningObjectTable.EnumRunning(out monikerList);
                monikerList.Reset();

                IMoniker[] monikerContainer = new IMoniker[1];
                IntPtr pointerFetchedMonikers = IntPtr.Zero;

                // fetch all moniker
                while (monikerList.Next(1, monikerContainer, pointerFetchedMonikers) == 0)
                {
                    // create binding object
                    IBindCtx bindInfo;
                    CreateBindCtx(0, out bindInfo);

                    // query com proxy info      
                    object comInstance = null;
                    runningObjectTable.GetObject(monikerContainer[0], out comInstance);

                    string ppszDisplayName;
                    try { monikerContainer[0].GetDisplayName(bindInfo, null, out ppszDisplayName); }
                    catch { ppszDisplayName = ""; }

                    Guid pClassID;
                    try { monikerContainer[0].GetClassID(out pClassID); }
                    catch { pClassID = Guid.Empty; }

                    string ClassName;
                    try { ClassName = TypeDescriptor.GetClassName(comInstance); }
                    catch { ClassName = ""; }

                    string ComponentName;
                    try { ComponentName = TypeDescriptor.GetComponentName(comInstance, false); }
                    catch { ComponentName = ""; }

                    if ((RunningComponent.DisplayName == ppszDisplayName) &&
                        (RunningComponent.ClsID == pClassID) &&
                        (RunningComponent.ComponentClassName == ClassName) &&
                        (RunningComponent.ComponentName == ComponentName)
                        )
                    {
                        Marshal.ReleaseComObject(bindInfo);
                        ResultInstance = comInstance;
                    }
                    else
                        Marshal.ReleaseComObject(comInstance);

                    Marshal.ReleaseComObject(bindInfo);
                }
                
                return ResultInstance;
            }
            finally
            {
                // release proxies
                if (runningObjectTable != null)
                    Marshal.ReleaseComObject(runningObjectTable);
                if (monikerList != null)
                    Marshal.ReleaseComObject(monikerList);
            }
        }

        //convert from System.Runtime.InteropServices.ComTypes.FILETIME to the local System.DateTime
        // The COM FILETIME is based on coordinated universal time (UTC). UTC-based time is loosely defined as the current date and time of day in Greenwich, England.
        // this function uses the current system settings for the time zone (differenc from UTC to your timezone) and daylight saving time (summer- / winter-time settings).
        static DateTime ConvertFromFILETIME(System.Runtime.InteropServices.ComTypes.FILETIME FileTime)
        {
            long highBits = FileTime.dwHighDateTime;
            highBits = highBits << 32;
            try { return DateTime.FromFileTimeUtc(highBits | (long)(uint)FileTime.dwLowDateTime).ToLocalTime(); }
            catch { return DateTime.MinValue; }
        }

        //convert from System.DateTime to System.Runtime.InteropServices.ComTypes.FILETIME
        static System.Runtime.InteropServices.ComTypes.FILETIME ConvertToFILETIME(DateTime TimeAndDate)
        {
            System.Runtime.InteropServices.ComTypes.FILETIME FileTime = new System.Runtime.InteropServices.ComTypes.FILETIME();
            long dtFileTime = TimeAndDate.ToFileTime();
            FileTime.dwLowDateTime = (int)(dtFileTime & 0xFFFFFFFF);
            FileTime.dwHighDateTime = (int)(dtFileTime >> 32);
            return FileTime;
        }
    } // end class RunningObjectTable

    // Simple class to create an unbound object to hold data out of an
    // Component of the Running Object Table (ROT) 
    public class RunningObjectTableComponentInfo
    {
        // constructor
        public RunningObjectTableComponentInfo(
            String DisplayName,
            Guid ClassID, DateTime LastChanged,
            long SizeMax,
            String ComponentName,
            String ComponentClassName,
            Boolean IsRunning,
            Boolean IsDirty
         )
        {
            _DisplayName = DisplayName;
            _ClsID = ClassID;
            _LastChanged = LastChanged;
            _SizeMax = SizeMax;
            _ComponentName = ComponentName;
            _ComponentClassName = ComponentClassName;
            _IsRunning = IsRunning;
            _IsDirty = IsDirty;
        }

#region properties
        private String _DisplayName;
        public String DisplayName
        {get { return _DisplayName; }}

        private Guid _ClsID;
        public Guid ClsID
        {get { return _ClsID; }}

        internal DateTime _LastChanged;
        public DateTime LastChanged
        {
            get {
                this.Refresh();
                return _LastChanged;
            }
        }

        private long _SizeMax;
        public long SizeMax
        {get { return _SizeMax; }}

        private String _ComponentName;
        public String ComponentName
        {get { return _ComponentName; }}

        private String _ComponentClassName;
        public String ComponentClassName
        {get { return _ComponentClassName; }}

        internal Boolean _IsRunning = false;
        public Boolean IsRunning
        {
            get {
                this.Refresh();
                return _IsRunning;
            }
        }

        internal Boolean _IsDirty = false;
        public Boolean IsDirty
        {
            get {
                this.Refresh();
                return _IsDirty;
            }
        }

#endregion // properties

        // function to refresh the property values of an RunningObjectTableComponentInfo object 
        private void Refresh()
        {
            // set default values 
            this._IsDirty = false;
            this._IsRunning = false;

            // refresh the properties of the component by a call to GetComponentsFromROT()
            foreach (RunningObjectTableComponentInfo RunningROTComponent in RunningObjectTable.GetComponentsFromROT())
            {
                // if the object is still in the ROT table we refresh the property values
                if (RunningROTComponent.DisplayName == this.DisplayName)
                {
                    // if the object is still in the ROT table and the refreshed property _IsRunning is still true we return true
                    this._IsDirty = RunningROTComponent._IsDirty;
                    // if the object is still in the ROT table and the refreshed property _IsRunning is still true we return true
                    this._IsRunning = RunningROTComponent._IsRunning;
                    // if the object is still in the ROT table we return the refreshed LastChanged DateTime
                    this._LastChanged = RunningROTComponent._LastChanged;
                    break;
                }
            }
        }
        
        // if this object is still in the Running Object Table (ROT) we try to return the reference to it
        public object GetInstance() {
            return RunningObjectTable.GetInstanceFromROT(this);
        }

    }  // end class RunningObjectTableComponentInfo

} // end namespace
'@

Add-Type -TypeDefinition $TypeDefinition

Function Get-ComRot {
<#
.Synopsis
  Gets a list of the COM objects out of the Running Object Table (ROT) which holds the current runnig COM Objects

.DESCRIPTION
  Gets a list of the COM objects out of the Running Object Table (ROT) which holds the current runnig COM Objects
  The Running Object Table (ROT) is a machine-wide table in which objects can register themselves.
  Because it is optional for a COM object to register within the ROT you can not find all running COM Objects inside this table
  Even sometimes you will find only one entry in the ROT for an COM object which runs many instances. This is by design of the COM producer.

.PARAMETER EnrichWithRegistry
  Use this parameter to find the ProgID out of the Registry for each resulting object.
  If this parameter is used, the Get-ComRotRegistered function adds the following ProgID information out of the registry to each result object:
  
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

.PARAMETER DoNotUseCache
  This is a dynamic parameter which is only showing up if the -EnrichWithRegistry parameter is used!
 	If you like use fresch registry informations out of the registry use the -DoNotUseCache parameter of the functions.
	If you use the -DoNotUseCache parameter, the cache is also freshed up.

  	The Registry entrys are changing only if a new COM Component ist registered or unregistered (or Sofware was installed).
	This happens not very often so this function take use of an array as a cache mechanism.
	The advantage is, that to read informations out of the cache is faster then reading it out of the registry every time this function is called.
	If you like use fresch registry informations out of the registry use the -DoNotUseCache parameter of the functions.
	If you use the -DoNotUseCache of any function, the cache is also freshed up.

.EXAMPLE
   Get-ComRot

   Get all COM objects out of the Running Object Table (Rot)

.EXAMPLE
   Get-ComRot -EnrichWithRegistry

   Get all COM objects out of the Running Object Table (Rot) with additional ProgID informations out of the registry.

.EXAMPLE
  $WordDoc = (Get-ComRot | Where-Object { $_.DisplayName -eq 'Document1'}).GetInstance()

  Get a moniker (a reference) to an Microsoft word document COM object with the DisplayName 'Document1' out of the Running Object Table (ROT)  

.OUTPUTS
  
  This function returns a List of Pamk_COM_ROT.RunningObjectTableComponentInfo objects.
  The Pamk_COM_ROT.RunningObjectTableComponentInfo is a custom .NET class and
  it is similar to the System.Runtime.InteropServices.ComTypes.IMoniker interface.
  The reason because we use a custom class instead of the IMoniker interface is,
  that the IMoniker objects are bound to the reference of the COM object and the RunningObjectTableComponentInfo are not.
  The RunningObjectTableComponentInfo objects are a dumb Property bag without bindig to associated the COM object.
  This unbound behavior is done to make it ease to destroy thec COM objects and not adding an additional referenc to the reference counter.
  So if your ROT query was a time ago, you can not be sure that an object out of this list is still alive.
  Use the RunningObjectTableComponentInfoobject.IsRunning property or make shure that you always work with a fresh query result.
   
  RunningObjectTableComponentInfoobject has the following properties:  
   
  DisplayName
    The display name, which is a user-readable representation of the current moniker.
      
  ClsID
    the class identifier (CLSID) of an object.

  LastChanged
    Provides a DateTime object representing the time that the object identified by the current moniker was last changed.
    (This property is not well maintained by the most COM classes!)

  SizeMax
    Returns the size, in bytes, of the stream needed to save the object.

  ComponentName
    The name of the class for the specified component.
    (obtained by use of the System.ComponentModel.Typedescriptor .NET class)

  ComponentClassName
    The name of the specified component.
    (obtained by use of the System.ComponentModel.Typedescriptor .NET class)

  IsRunning
    Determines whether the object that is identified by the current moniker is currently loaded and running.
    (This property is always doing a fresh request to the Running Object Table so it is a reliable source) 

  IsDirty
    Checks the object for changes since it was last saved.
    (In my test this property is never changing from the COM objects, and was allways $false!)

    RunningObjectTableComponentInfoobject has the following methods:
      RunningObjectTableComponentInfoobject.GetInstance()
        Use this method to obtain a reference to the instance of a COM object out of the Running Object Table (ROT)
        The returntype is System.Object and holds an reference to a COM object which in fact is a type of System.Runtime.InteropServices.ComTypes.IMoniker.
        So this method is similar to the Visual Basic method GetObject(). 

.NOTES
  Author: Peter Kriegel
  Version 1.0.0. 7.November.2014
#>

  [CmdletBinding()]
  param(
    [Switch]$EnrichWithRegistry
  )

  DynamicParam {
    if ($EnrichWithRegistry.IsPresent) {
      
        # the -DoNotUseCache parameter makes only sense if the -EnrichWithRegistry parameter is set
        # we can not use parameterset here because we have only one parameter, so we have to create a dynamic parameter 

        # create attribute to set the parameter in the default parameterset with the name '__AllParameterSets'
        $Attributes = New-Object 'Management.Automation.ParameterAttribute'
        $Attributes.ParameterSetName = '__AllParameterSets'
        $AttributesCollection = New-Object 'Collections.ObjectModel.Collection[Attribute]'
        $AttributesCollection.Add($Attributes)
        # create the parameter
        $DoNotUseCache = New-Object 'System.Management.Automation.RuntimeDefinedParameter' -ArgumentList @('DoNotUseCache',[Switch],$AttributesCollection)
        # add parameter to dictionary
        $paramDictionary = New-Object 'System.Management.Automation.RuntimeDefinedParameterDictionary'
        $paramDictionary.Add('DoNotUseCache', $DoNotUseCache)
        # return dictionary which contains the dynamic parameters
        # it is added to the $PSBoundParameters dictionary of the function
        return $paramDictionary
    }
  } # end of DynamicParam block
  
  end {
        
    $Result = [Pamk_COM_ROT.RunningObjectTable]::GetComponentsFromROT()
  
    IF($EnrichWithRegistry.IsPresent) {
      ForEach($RunningObjectTableComponentInfo in $Result) {
        Get-ComRotRegistered -RunningObjectTableComponentInfo $RunningObjectTableComponentInfo -DoNotUseCache:([bool]($PSBoundParameters['DoNotUseCache']).IsPresent)
      }
    } Else {
      $Result
    }
  } # end of end block
}
Function Get-ComRotIntance {
<#
.Synopsis
   Function to obtain a reference to the instance of a COM object out of the Running Object Table (ROT)

.DESCRIPTION
  Function to obtain a reference to the instance of a COM object out of the Running Object Table (ROT)
  So this method is similar to the Visual Basic method GetObject().

.EXAMPLE
  $Obj = Get-ComRotIntance -RunningObjectTableComponentInfo (Get-ComRot | Where-Object { $_.ComponentName -eq 'Microsoft Excel' } | Select-Object -First 1)
  
  $Obj.Visible = $True
  
  # get first Workbook from Excel Object
  $workbook=$Obj.Workbooks.Item(1)
  
  # adding a new worksheet to the Worbook
  $worksheet = $workbook.Worksheets.Item(1)
  
  # filling some cells with use of Range Object
  $worksheet.Range('A7').Value2 = 7
  $worksheet.Range('A1').Value2 = 1
  $worksheet.Range('A2').Value2 = 2
  $worksheet.Range('A3').Value2 = 3
  $worksheet.Range('A4').Value2 = 4

  Get Excel instance out of the Running Object table (ROT) and wor with it.

.OUTPUTS
  The returntype is System.Object and holds an reference to a COM object which in fact is a type of System.Runtime.InteropServices.ComTypes.IMoniker.
  If the object was removed from the Running Object Table (ROT) the returntype is $Null

.NOTES
  Author: Peter Kriegel
  Version 1.0.0. 7.November.2014

#>

  Param(
    [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
    [Pamk_COM_ROT.RunningObjectTableComponentInfo]$RunningObjectTableComponentInfo
  )

  [Pamk_COM_ROT.RunningObjectTable]::GetInstanceFromROT($RunningObjectTableComponentInfo)

}