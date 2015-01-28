using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using UIAutomationClient;
using System.Diagnostics;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Collections;

namespace GuiGood
{
    public class FunctionsLibrary
    {
        //Variables
        GuiGood gui;
        public string ProcessName = "Systest";

        //GUI Functions
        /// <summary>
        /// Set the GUI reference
        /// </summary>
        /// <param name="gui1"></param>
        public void SetGUI(GuiGood gui1)
        {
            gui = gui1;
        }
        /// <summary>
        /// Load Objects into ListView
        /// </summary>
        public void LoadObjects()
        {
            //Grab Process
            Process[] processes = Process.GetProcessesByName(ProcessName);
            //Foreach process
            foreach (Process p in processes)
            {
                //Declare Variables
                AutomationElement window = AutomationElement.FromHandle(p.MainWindowHandle);
                AutomationElementCollection listofelements = FindAllEnabled(window);
                
               // IEnumerable<AutomationElementCollection> myEnumerable = new IEnumerable<AutomationElementCollection>(listofelements);
                foreach (AutomationElement elem in listofelements)
                {
                    if (elem.Current.Name != "")
                    {
                        gui.AddListView(elem.Current.Name, elem.Current.LocalizedControlType.ToString());
                    }
                }
                gui.AddElementList(listofelements);
            }
        }

        //Events
        /// <summary>
        /// Set Event Handlers
        /// </summary>
        public void SetEventHandler()
        {
            //Grab Process
            Process[] processes = Process.GetProcessesByName(ProcessName);
            //Foreach process
            foreach (Process p in processes)
            {
                //Declare Variables
                AutomationElement window = AutomationElement.FromHandle(p.MainWindowHandle);
            Automation.AddStructureChangedEventHandler(window, System.Windows.Automation.TreeScope.Children, new StructureChangedEventHandler(OnStructureChanged));
            }
        }
        /// <summary>
        /// Event Handler for Structure Changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnStructureChanged(object sender, StructureChangedEventArgs e)
        {
            gui.ClearColumn();
            LoadObjects();
            gui.SortColumn();
        }

        //UI Automation
        /// <summary>
        /// Find all Enabled Elements in Process Window
        /// </summary>
        /// <param name="elementWindowElement"></param>
        /// <returns></returns>
        public AutomationElementCollection FindAllEnabled(AutomationElement elementWindowElement)
        {
            if (elementWindowElement == null)
            {
                throw new ArgumentException();
            }
            Condition conditions = new PropertyCondition(AutomationElement.IsEnabledProperty, true);

            // Find all children that match the specified conditions.
            AutomationElementCollection elementCollection =
                elementWindowElement.FindAll(System.Windows.Automation.TreeScope.Descendants, conditions);
            return elementCollection;
        }

    }
}
