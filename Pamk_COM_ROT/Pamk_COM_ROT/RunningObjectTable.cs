using System;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.ComponentModel;
using System.Collections.Generic;
using System.Security.Permissions;
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
} // end namespace