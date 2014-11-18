using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace COM_ROT
{
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

        // function to refresh the property values of an RunningObjectTableComponent object 
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

    }  // end class RunningObjectTableComponent
}
