using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;

namespace GuiGood
{
    class AutomationFunctions
    {
        //Declerations
        #region Decleration
        private const int BM_CLICK = 0x00F5;
        private const uint WM_GETTEXT = 0x000D;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);
        private const UInt32 MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const UInt32 MOUSEEVENTF_LEFTUP = 0x0004;
        [DllImport("user32.dll")]
        private static extern void mouse_event(UInt32 dwFlags, UInt32 dx, UInt32 dy, UInt32 dwData, IntPtr dwExtraInfo);
        #endregion

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

        /// <summary>
        /// Click a button from a name or string
        /// </summary>
        /// <param name="name">Name of the element</param>
        public bool ClickButtonWithName(string name, string ProcessName)
        {
            //Grab Process
            Process[] processes = Process.GetProcessesByName(ProcessName);
            //Foreach process
            foreach (Process p in processes)
            {
                //Declare Variables
                AutomationElement window = AutomationElement.FromHandle(p.MainWindowHandle);
                AutomationElementCollection listofelements = FindAllEnabled(window);
                //Set Focus on the Window
                try
                {
                    //SetFocusScreen();
                }
                catch { }
                //Iterate Elements
                foreach (AutomationElement elem in listofelements)
                {
                    if (elem.Current.ControlType == ControlType.Button && elem.Current.Name == name)
                    {
                        ClickButton(elem, "button");
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Set Focus on Main Window
        /// </summary>
        /// <param name="ProcessName"></param>
        public void SetFocusMainWindow(string ProcessName)
        {
            //Grab Process
            Process[] processes = Process.GetProcessesByName(ProcessName);
            //Foreach process
            foreach (Process p in processes)
            {
                //Declare Variables
                AutomationElement window = AutomationElement.FromHandle(p.MainWindowHandle);
                try
                {
                    window.SetFocus();
                }
                catch { }
            }
        }

        /// <summary>
        /// Insert Text to textbox
        /// </summary>
        /// <param name="name"></param>
        /// <param name="ProcessName"></param>
        /// <param name="text"></param>
        public bool InsertText(string name, string ProcessName, string text)
        {
            //Grab Process
            Process[] processes = Process.GetProcessesByName(ProcessName);
            //Foreach process
            foreach (Process p in processes)
            {
                //Declare Variables
                AutomationElement window = AutomationElement.FromHandle(p.MainWindowHandle);
                AutomationElementCollection listofelements = FindAllEnabled(window);
                //Set Focus on the Window
                try
                {
                    //SetFocusScreen();
                }
                catch { }
                //Iterate Elements
                foreach (AutomationElement elem in listofelements)
                {
                    if (elem.Current.ControlType == ControlType.Edit && elem.Current.Name == name)
                    {
                        InsertTextUsingUIAutomation(elem, text);
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Click Panel with name
        /// </summary>
        /// <param name="name"></param>
        /// <param name="ProcessName"></param>
        public bool ClickPanelWithName(string name, string ProcessName)
        {
            //Grab Process
            Process[] processes = Process.GetProcessesByName(ProcessName);
            //Foreach process
            foreach (Process p in processes)
            {
                //Declare Variables
                AutomationElement window = AutomationElement.FromHandle(p.MainWindowHandle);
                AutomationElementCollection listofelements = FindAllEnabled(window);
                //Set Focus on the Window
                try
                {
                    //SetFocusScreen();
                }
                catch { }
                //Iterate Elements
                foreach (AutomationElement elem in listofelements)
                {
                    if (elem.Current.ControlType == ControlType.Pane && elem.Current.Name == name)
                    {
                        ClickPanel(elem);
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Check if page contains element
        /// </summary>
        /// <param name="name"></param>
        /// <param name="ProcessName"></param>
        public bool doesPageContainElement(string name, string ProcessName)
        {
            //Grab Process
            Process[] processes = Process.GetProcessesByName(ProcessName);
            //Foreach process
            foreach (Process p in processes)
            {
                //Declare Variables
                AutomationElement window = AutomationElement.FromHandle(p.MainWindowHandle);
                AutomationElementCollection listofelements = FindAllEnabled(window);
                //Iterate Elements
                foreach (AutomationElement elem in listofelements)
                {
                    try
                    {
                        if (elem.Current.Name == name)
                        {
                            return true;
                        }
                    }
                    catch
                    {
                        return false;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Click a Button Object through Send Message
        /// </summary>
        /// <param name="handle"></param>
        public void ClickButton(AutomationElement elem, string name)
        {
            //Click Button
            //SendMessage(handle, (uint)BM_CLICK, IntPtr.Zero, IntPtr.Zero);
            //Automation Click
            // var buttonWindow = AutomationElement.RootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
            // if (buttonWindow == null) return;
            var invokePattern = elem.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern.Invoke();

        }

        /// <summary>
        /// Click a panel from element
        /// </summary>
        /// <param name="elem">element in which to click</param>
        public void ClickPanel(AutomationElement elem)
        {
            Cursor.Position = new System.Drawing.Point((int)elem.GetClickablePoint().X, (int)elem.GetClickablePoint().Y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, new IntPtr());
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, new IntPtr());
        }

        /// <summary>
        /// Set Item in ComboBox
        /// </summary>
        /// <param name="name"></param>
        /// <param name="ProcessName"></param>
        /// <param name="text"></param>
        public bool SetItem(string name, string ProcessName, string text)
        {
            //Grab Process
            Process[] processes = Process.GetProcessesByName(ProcessName);
            //Foreach process
            foreach (Process p in processes)
            {
                //Declare Variables
                AutomationElement window = AutomationElement.FromHandle(p.MainWindowHandle);
                AutomationElementCollection listofelements = FindAllEnabled(window);
                //Set Focus on the Window
                try
                {
                    //SetFocusScreen();
                }
                catch { }
                //Iterate Elements
                foreach (AutomationElement elem in listofelements)
                {
                    if (elem.Current.ControlType == ControlType.ComboBox && elem.Current.Name == name)
                    {
                        SetSelectedComboBoxItem(elem, text);
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Get Selected comboBox
        /// </summary>
        /// <param name="comboBox">comboBox element in which you want to select an item from</param>
        /// <param name="item">name of the item you want to select</param>
        public static void SetSelectedComboBoxItem(AutomationElement comboBox, string item)
        {
            AutomationPattern automationPatternFromElement = GetSpecifiedPattern(comboBox, "ExpandCollapsePatternIdentifiers.Pattern");

            ExpandCollapsePattern expandCollapsePattern = comboBox.GetCurrentPattern(automationPatternFromElement) as ExpandCollapsePattern;

            expandCollapsePattern.Expand();
            expandCollapsePattern.Collapse();

            AutomationElement listItem = comboBox.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.NameProperty, item));

            automationPatternFromElement = GetSpecifiedPattern(listItem, "SelectionItemPatternIdentifiers.Pattern");

            SelectionItemPattern selectionItemPattern = listItem.GetCurrentPattern(automationPatternFromElement) as SelectionItemPattern;
            selectionItemPattern.Select();
        }

        /// <summary>
        /// Get the patterns for combo box
        /// </summary>
        /// <param name="element">comboBox element to get patterns from</param>
        /// <param name="patternName">the specific Pattern you want to get</param>
        /// <returns></returns>
        public static AutomationPattern GetSpecifiedPattern(AutomationElement element, string patternName)
        {
            AutomationPattern[] supportedPattern = element.GetSupportedPatterns();

            foreach (AutomationPattern pattern in supportedPattern)
            {
                // if (pattern.ProgrammaticName.ToLower().Contains(patternName))
                if (pattern.ProgrammaticName == patternName)
                    return pattern;
            }

            return null;
        }

        /// <summary>
        /// Insert text into a textbox element
        /// </summary>
        /// <param name="element">the textbox Element</param>
        /// <param name="value">the value in which to insert</param>
        public void InsertTextUsingUIAutomation(AutomationElement element, string value)
        {
            try
            {
                // Validate arguments / initial setup 
                if (value == null)
                    throw new ArgumentNullException(
                        "String parameter must not be null.");

                if (element == null)
                    throw new ArgumentNullException(
                        "AutomationElement parameter must not be null");

                // A series of basic checks prior to attempting an insertion. 
                // 
                // Check #1: Is control enabled? 
                // An alternative to testing for static or read-only controls  
                // is to filter using  
                // PropertyCondition(AutomationElement.IsEnabledProperty, true)  
                // and exclude all read-only text controls from the collection. 
                if (!element.Current.IsEnabled)
                {
                    throw new InvalidOperationException(
                        "The control with an AutomationID of "
                        + element.Current.AutomationId.ToString()
                        + " is not enabled.\n\n");
                }

                // Check #2: Are there styles that prohibit us  
                //           from sending text to this control? 
                if (!element.Current.IsKeyboardFocusable)
                {
                    throw new InvalidOperationException(
                        "The control with an AutomationID of "
                        + element.Current.AutomationId.ToString()
                        + "is read-only.\n\n");
                }


                // Once you have an instance of an AutomationElement,   
                // check if it supports the ValuePattern pattern. 
                object valuePattern = null;

                // Control does not support the ValuePattern pattern  
                // so use keyboard input to insert content. 
                // 
                // NOTE: Elements that support TextPattern  
                //       do not support ValuePattern and TextPattern 
                //       does not support setting the text of  
                //       multi-line edit or document controls. 
                //       For this reason, text input must be simulated 
                //       using one of the following methods. 
                //        
                if (!element.TryGetCurrentPattern(
                    ValuePattern.Pattern, out valuePattern))
                {


                    // Set focus for input functionality and begin.
                    element.SetFocus();

                    // Pause before sending keyboard input.
                    Thread.Sleep(100);

                    // Delete existing content in the control and insert new content.
                    SendKeys.SendWait("^{HOME}");   // Move to start of control
                    SendKeys.SendWait("^+{END}");   // Select everything
                    SendKeys.SendWait("{DEL}");     // Delete selection
                    SendKeys.SendWait(value);
                }
                // Control supports the ValuePattern pattern so we can  
                // use the SetValue method to insert content. 
                else
                {


                    // Set focus for input functionality and begin.
                    element.SetFocus();

                    ((ValuePattern)valuePattern).SetValue(value);
                }
            }
            catch (ArgumentNullException exc)
            {

            }
            catch (InvalidOperationException exc)
            {

            }
            finally
            {

            }
        }

        /// <summary>
        /// Write the MMF
        /// </summary>
        /// <param name="info">the obd string to write</param>
        public MemoryMappedFile writeMMF(string name, string info)
        {
            byte[] buffer = Encoding.UTF8.GetBytes(info);
            var mmf = MemoryMappedFile.CreateOrOpen(name, 1000);
            var writer = mmf.CreateViewAccessor();
            writer.Write(54, (ushort)buffer.Length);
            writer.WriteArray(54 + 2, buffer, 0, buffer.Length);
            GC.KeepAlive(mmf);
            return mmf;
        }

        /// <summary>
        /// Close Memory Mapped File
        /// </summary>
        /// <param name="mmf"></param>
        public void closeMMF(MemoryMappedFile mmf)
        {
            mmf.Dispose();
        }
    }
}
