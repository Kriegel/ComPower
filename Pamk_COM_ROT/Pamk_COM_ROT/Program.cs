using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pamk_COM_ROT
{
    class Program
    {
        static void Main(string[] args)
        {
    

            // get all components from the ROT table
            IList<RunningObjectTableComponentInfo> ROTComponents = RunningObjectTable.GetComponentsFromROT();
            
            Console.WriteLine("Component count: " + ROTComponents.Count.ToString());

            // display all properties out of the found components
            foreach (RunningObjectTableComponentInfo ROTComponent in ROTComponents)
            {
                Console.WriteLine("Displayname: " + ROTComponent.DisplayName);
                Console.WriteLine("ClassID: " + ROTComponent.ClsID);
                Console.WriteLine("LastChanged: " + ROTComponent.LastChanged);
                Console.WriteLine("SizeMax: " + ROTComponent.SizeMax);
                Console.WriteLine("ComponentName: " + ROTComponent.ComponentName);
                Console.WriteLine("ComponentClassName: " + ROTComponent.ComponentClassName);
                Console.WriteLine("Component is running ? " + ROTComponent.IsRunning.ToString());
                Console.WriteLine("Component is dirty ? " + ROTComponent.IsDirty.ToString());
                Console.WriteLine("");

                object comInstance = ROTComponent.GetInstance();
                System.Threading.Thread.Sleep(500);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(comInstance);
                System.Threading.Thread.Sleep(500);
            }
            Console.ReadLine();
        }
    }
}