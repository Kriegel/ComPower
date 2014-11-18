Function Remove-ComObject { 
<# 
.Synopsis 
     Releases a given Variable which holds a reference to an COM object or relaese all <__ComObject> objects in the caller scope.
     (This function utilize a call to System.Runtime.Interopservices.Marshal]::ReleaseComObject() see warning inside the description!) 
      
.Description 
     Releases a given Variable which holds a reference to an COM object or relaese all <__ComObject> objects in the caller scope.
     
     Because the COM objects live in unmanaged memory and Windows PowerShell lives in managed .NET memory the binding to each other is VERY fragile!
     So it is very hard to clean up a COM object from .NET (Windows PowerShell).

     Spirits that I've cited
     My commands ignore.
      (Johann Wolfgang von Goethe)

     There are millons of discussions in the internet how to clean up COM objects correctly from .NET.
     For Microsoft office applications (and other) follow these Steps (using Excel as example here):
     
     1. Create an explicit PowerShell variable for every Excel object you use, even for every sub-object, to hold a reference to the object.
        The main object (Excel.Application) holds references (in his reference counter) to each sub-object,
        and the main object can not be closed until all sub objects are closed, and the reference counter count is 0.  
        
        # get a new instance of Excel as COM-Object, and hold reference in variable $excel
        $excel = new-object -comobject excel.application
        $excel.visible = $true

        # Adding a new Workbook to Excel Object, and hold reference in variable $workbook
        $workbook = $excel.Workbooks.add() # add reference to Excel.Application reference counter
 
     
     2. Close all COM object over their close() or quit() methods if they have some!
        You have to walk up from bottom to top in the object hirarchy to close objects, so close sub-sub-sub object first and walk up.
        (Attention: the wording inside COM is very inconsistent! SO you have to use the Close() and Quit() methods here.)
          
          $workbook.Close($False)
          $excel.Quit()
    
     3. Do a call of [System.Runtime.Interopservices.Marshal]::ReleaseComObject(<Variable>) for each PowerShel variable which holds a reference to a COM object
        (this step is done by this function Remove-ComObject)
        
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

        WARNING!
        The call to [System.Runtime.Interopservices.Marshal]::ReleaseComObject() will free ALL references to the COM object.
        So if the COM object is used on another place in PowerShell, the COM object can not be used anywhere after a call to ReleaseComObject() .
        See blog post: ReleaseComObject 
        http://blogs.msdn.com/b/cbrumme/archive/2003/04/16/51355.aspx
                
     4. Remove the now empty variable from the PowerShell variable: drive by use of the cmdlet Remove-Variable
        (this step is done by this function Remove-ComObject)
        
        Remove-Variable -Name workbook
        Remove-Variable -Name excel
        
     5. Optional free up the used memory to kill the refrences to the unmanaged memory definitely with [System.GC]::Collect() and [System.GC]::WaitForPendingFinalizers()
        (this step is done by this function Remove-ComObject)

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
                    
     Sometimes even the Voodo inside of this function don't help or can't be appropriate!
     Then you have only the following additional options:

     Try to send a WM_CLOSE message to the main window of the Process which holds the COM Objects.
     If the Process which holds the COM object is still alive, you have to kill the process to get rid of the COM objects
     (Kill process by the his main MainWindowHandle or MainWindowTitle of its Window or do this : Get-Process *Excel* | Stop-Process)
     
     You can even use create a "job object" for your main application.
     Then configure the job object to kill-process upon close.
     If the application dies, the OS will take care of cleaning up your processes.
     See:
     Closing an Excel Interop process cleanly, even if your application crashes
     http://www.xtremevbtalk.com/showpost.php?p=1335552&postcount=22

     For microsoft Office see KB317109 :
     http://support2.microsoft.com/default.aspx?scid=kb;EN-US;Q317109

     Good general discussions about cleaning COM objects you can find here:
     How to properly clean up Excel (Office) interop objects
     http://stackoverflow.com/questions/158706/how-to-properly-clean-up-excel-interop-objects

     Does *every* Excel interop object need to be released using Marshal.ReleaseComObject?
     http://stackoverflow.com/questions/2926205/does-every-excel-interop-object-need-to-be-released-using-marshal-releasecomob

     Find references to the COM object in runtime
     http://stackoverflow.com/questions/1471927/find-references-to-the-object-in-runtime

.PARAMETER Name
  Specifies the name of the variable which hold a reference to a COM to be removed.
  If the name is given the variable is searched up in the parent scopes beginning from the local scope by number of the scope.
  The first variable which matches the name and is holding a reference to an COM object in the parent scopes is removed.
  For more information, see about_Scopes.
  
  The parameter name ("Name") is optional.
  If no Name is given ALL variables which hold a reference to a COM object in the parent scope are Removed!

.PARAMETER DontCallReleaseComObject
  Use this parameter to supress a call to [System.Runtime.Interopservices.Marshal]::ReleaseComObject() 

.Example
     Remove-ComObject -Name 'ExcelApp' -Verbose 

     Releases the variable with the name 'ExcelApp' from the caller scope which has a Type of 'System.__ComObject' and displays the released COM objects' variable names.

.Example
     Remove-ComObject -Verbose 

     Releases all variables which has a Type of 'System.__ComObject' from the caller scope and displays the released COM objects' variable names. 

.Outputs 
     None 

.Notes 
     Orinator Author:    Robert Robelo
     https://gallery.technet.microsoft.com/office/d16d0c29-78a0-4d8d-9014-d66d57f51f63
     
     Improved by Author: Peter Kriegel
     First release:  01/13/2010 19:14
     Current Version: 2.0.0 from: 10.November.2014
      
 #> 
    [CmdletBinding()] 
    param(
        [String]$Name = '',
        [Switch]$DontCallReleaseComObject
    ) 

  begin {
  
    Function Get-ScopeDistanceToVariable {
      # Finds all Variables with a given name and
      # returns the scope distance numbers from this scope for each
      # found variable with this name 

        param(
          [String[]]$VariableName,
          [Switch]$FirstOnly
        )

        $i = 0
        # create endless loop
        While($True) { 
            Try {
              $Var = $Null
              $Var = Get-Variable -Name $VariableName -Scope $i -ErrorAction Stop
            } catch [System.Management.Automation.PSArgumentOutOfRangeException] {
              #reached the Global Scope!
              break
            } Catch { <# supress Errors#> }
            # If Variable with this Name exist inside the scope level return depth of scope number 
            If($Var) {
              Write-Output ($i - 1)
              If($FirstOnly.IsPresent) {break}
            }
            $i++
        }
      } # end of function Get-ScopeDistanceToVariable

  } # end of begin block

  process {

      # create ScopedItemOptions object to exclude ReadOnly and Constant objects from removing
      [Management.Automation.ScopedItemOptions]$scopedOpt = ([System.Management.Automation.ScopedItemOptions]::ReadOnly -bor [System.Management.Automation.ScopedItemOptions]::Constant)


       # if VariableName was given we process only this variable with this name
       If(-not [String]::IsNullOrEmpty($Name)) {
          
          ForEach($Varname in $Name) {
          
            # try to clean up a single variable which holds a COM Object
        
            # try to get the Variable from parent scope
            #$Var = 

            $ScopeNumber = Get-ScopeDistanceToVariable -VariableName $Varname -FirstOnly

            $Var = Get-Variable -Name $Varname -Scope $ScopeNumber
        
            # test if the variable value holds a reference to an COM object if not abort operation here
            If(-Not $Var.Value.GetType().Name -like '*__ComObject*' ) {
                Write-Warning "Variable with Name: $Varname seems not to hold a COM object. I Abort!"
                break
            }

            # using .NET method to decrement the reference counter of the Runtime Callable Wrapper (RCW) associated with the specified COM object
            # This method is used to explicitly control the lifetime of a COM object used from managed code.
            # You should use this method to free the underlying COM object that holds references to resources in a timely manner or
            # when objects must be freed in a specific order.
            # Every time a COM interface pointer enters the common language runtime (CLR), it is wrapped in an RCW.
            # The RCW has a reference count that is incremented every time a COM interface pointer is mapped to it.
            # The ReleaseComObject method decrements the reference count of an RCW. When the reference count reaches zero,
            # the runtime releases all its references on the unmanaged COM object, and throws a System.NullReferenceException if
            # you attempt to use the object further. If the same COM interface is passed more than one time from unmanaged to managed code,
            # the reference count on the wrapper is incremented every time, and calling ReleaseComObject returns the number of remaining references.
            # This method enables you to force an RCW reference count release so that it occurs precisely when you want it to.
            # However, improper use of ReleaseComObject may cause your application to fail, or may cause an access violation.
            If(-not $DontCallReleaseComObject.IsPresent) {
              Write-Verbose "Calling  ReleaseComObject to variable: $($Var.Name)"
              [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Var.value)        
            }
            $Var = $Null
            Write-Verbose "Removing variable: '$Varname' from scope distance: $ScopeNumber"
            Remove-Variable -Name $Varname -Scope $ScopeNumber -Force -Verbose:$Verbose.IsPresent
          }
       } else {
        
          # VariableName was not given we process all reachable variables which hold a reference to COM object
          # documentation see stepps above
          # because we have no Variable Name we cannot determine wich Scope level ist targeted
          # so we use the Scope level of 2 which is the direct scope level outside this module function
         
           Get-Variable -Scope 2 | Where-Object { $_.Value } | Where-Object { (-not ($_.Options -band $scopedOpt)) -and ($_.Value.GetType().Name -like '*__ComObject*' ) } |
               ForEach-Object {

                       If(-not $DontCallReleaseComObject.IsPresent) {
                        Write-Verbose "Calling  ReleaseComObject to variable: $($_.Name)"
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($_.value)
                       }
                       Write-Verbose "Removing variable: $($_.Name)"
                       $_ | Remove-Variable -Scope 2 -Verbose:$Verbose.IsPresent
           } 
       }
  } # end of process block

  end {
      # pauses for half a second to let Windows PowerShell mark objects as Disposed and
      # ready to be collected by the Garbage Collector.
      Start-Sleep -Milliseconds 500

      # Calling GC.Collect() and GC.WaitForPendingFinalizers() to make CLR release unused com-objects
      # In most code examples you'll see for cleaning up COM objects from .NET,
      # the GC.Collect() and GC.WaitForPendingFinalizers() calls are made TWICE!
      # This should not be required, however, unless you are using Visual Studio Tools for Office (VSTO),
      # which uses finalizers that cause an entire graph of objects to be promoted in the finalization queue.
      # see discussion: http://stackoverflow.com/questions/158706/how-to-properly-clean-up-excel-interop-objects
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()

  } # end of end block

}