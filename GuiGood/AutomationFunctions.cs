using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace GuiGood
{
    class AutomationFunctions
    {
       

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
