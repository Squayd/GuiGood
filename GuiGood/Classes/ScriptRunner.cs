using System;
using System.Collections.Generic;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GuiGood
{
    public static class ScriptRunner
    {
        public static int waitTime;
        public static string scriptType;
        public static string ProjectName;
        public static string ProjectPath;
        public static string ScriptName;
        private static MemoryMappedFile mmf;
        private static string processName;
        private static bool hasEnded;


        public static void startScript()
        {
            switch (scriptType)
            {
                case "Chunking":
                    startChunkScript();
                    break;
                case "Default":
                    startDefaultScript();
                    break;
            }
        }


        public static bool startChunkScript()
        {
           if (ProjectPath != null && ProjectName != null && waitTime != null)
            {
                AutomationFunctions autoFunc = new AutomationFunctions();
                string[] invokeArray;
                string invokeText;
                string name;
                string[] attri;
                string attri1;
                string attri2;
                string attri3;
                List<string> pageIdentifiers = new List<string>();
                List<string> objectIdentifiers = new List<string>();
                int pageCount = 0;
                int pageCountDo = 1;
                
                //Get Page Identifiers
                using (StreamReader sw = new StreamReader(ProjectPath + "\\" + ScriptName))
                {
                    string line;
                    while ((line = sw.ReadLine()) != null)
                    {
                        if (!line.StartsWith("//"))
                        {
                            string[] parsed = line.Split('|');
                            name = parsed[0];
                            attri = parsed[2].Split('-');
                            //Is Page Identifier?
                            string pageIdentifier = attri[0].Substring(2, attri[0].Length - 2);
                            if (pageIdentifier == "Y")
                            {
                                pageIdentifiers.Add(line);
                                pageCount++;

                            }
                            else
                            {
                                objectIdentifiers.Add("|" + pageCount.ToString() + "|"  + line);
                            }
                        }
                    }
                }
                hasEnded = false;
                while (hasEnded == false) { 
                foreach (string page in pageIdentifiers)
                {  
                    string[] parsed = page.Split('|');
                    name = parsed[0];
                    attri = parsed[2].Split('-');
                    processName = attri[2].Substring(2, attri[2].Length - 2);
                    //Page Identifier
                    if (autoFunc.doesPageContainElement(name, processName))
                    {
                        //Do rest of path untill next Identifier
                        foreach (string line in objectIdentifiers)
                        {
                            if (line.Contains("|" + pageCountDo + "|"))
                            {
                                string command = line.Substring(3, line.Length - 3);
                                DoChunk(command, autoFunc);
                                

                            }
                        }
                    }
                    pageCountDo++;
                    }
                    pageCountDo = 1;
                }
                    
            }
               return true;
        }

        private static void DoChunk(string line, AutomationFunctions autoFunc)
        {
            string[] invokeArray;
            string invokeText;
            string name;
            string[] attri;
            string attri1;
            string attri2;
            string attri3;

            string[] parsed = line.Split('|');
            name = parsed[0];

            attri = parsed[2].Split('-');
            try
            {
                processName = attri[2].Substring(2, attri[2].Length - 2);
            }
            catch { }
            autoFunc.SetFocusMainWindow(processName);
            if (parsed[1].Contains("Button Press"))
            {
                try
                {
                    autoFunc.ClickButtonWithName(name, processName);
                    
                }
                catch { }
            }
            else
            if (parsed[1].Contains("Click Pane"))
            {
                try
                {
                    autoFunc.ClickPanelWithName(name, processName);
                }
                catch { }
            }
            else
            if (parsed[1].Contains("Set Text"))
            {
                try
                {
                    invokeArray = parsed[1].Split(':');
                    invokeText = invokeArray[1];
                    autoFunc.InsertText(name, processName, invokeText);
                }
                catch { }
            }
            else
            if (parsed[1].Contains("Select Item"))
            {
                try
                {
                    invokeArray = parsed[1].Split(':');
                    invokeText = invokeArray[1];
                    autoFunc.SetItem(name, processName, invokeText);
                }
                catch { }
            }
            else
            if (parsed[1].Contains("Set Focus"))
            {
                 //NOT WORKING YET
            }
            else
            if (parsed[1].Contains("Write Memory Mapped File"))
            {
                try
                {
                    string mmfName = parsed[2].Split(':')[0];
                    string mmfDescrip = parsed[2].Split(':')[1];
                    mmf = autoFunc.writeMMF(mmfName, mmfDescrip);
                }
                catch { }
            }
           else
           if (parsed[1].Contains("Close Memory Mapped File"))
           {
               try
               {
                   autoFunc.closeMMF(mmf);
               }
               catch { }
           }
           else
           if (parsed[1].Contains("Send Message"))
           {
               try
               {
                   SendKeys.SendWait(parsed[2]);
               }
               catch { }
           }
           else
           if (parsed[1].Contains("Sleep"))
           {
               try
               {
                   Thread.Sleep(Convert.ToInt32(parsed[2]));
               }
               catch { }
           } else
           if (parsed[1].Contains("END"))
           {
               hasEnded = true;

           }
        }
        
        public static void startDefaultScript()
        {
            if (ProjectPath != null && ProjectName != null && waitTime != null)
            {
                AutomationFunctions autoFunc = new AutomationFunctions();
                string[] invokeArray;
                string invokeText;
                string name;
                string[] attri;
                string attri1;
                string attri2;
                string attri3;
                
                //Start Script
                using (StreamReader sw = new StreamReader(ProjectPath + "\\" + ScriptName))
                {
                    string line;
                    while ((line = sw.ReadLine()) != null)
                    {
                        if (!line.StartsWith("//"))
                        {
                            string[] parsed = line.Split('|');
                            name = parsed[0];

                            attri = parsed[2].Split('-');
                            try
                            {
                                processName = attri[2].Substring(2, attri[2].Length - 2);
                            }
                            catch { }
                            autoFunc.SetFocusMainWindow(processName);
                            if (parsed[1].Contains("Button Press"))
                            {
                                autoFunc.ClickButtonWithName(name, processName);
                            } else
                            if (parsed[1].Contains("Click Pane"))
                            {
                                autoFunc.ClickPanelWithName(name, processName);
                            } else 
                            if (parsed[1].Contains("Set Text"))
                            {
                                invokeArray = parsed[1].Split(':');
                                invokeText = invokeArray[1];
                                autoFunc.InsertText(name, processName, invokeText);
                            } else
                            if (parsed[1].Contains("Select Item"))
                            {
                                invokeArray = parsed[1].Split(':');
                                invokeText = invokeArray[1];
                                autoFunc.SetItem(name, processName, invokeText);  
                            } else
                            if (parsed[1].Contains("Set Focus"))
                            {
                                //NOT WORKING YET
                            } else
                            if (parsed[1].Contains("Write Memory Mapped File"))
                            {
                                string mmfName = parsed[2].Split(':')[0];
                                string mmfDescrip = parsed[2].Split(':')[1];
                                mmf = autoFunc.writeMMF(mmfName, mmfDescrip);
                            } else
                            if (parsed[1].Contains("Close Memory Mapped File"))
                            {
                                autoFunc.closeMMF(mmf);
                            } else
                            if (parsed[1].Contains("Send Message"))
                            {
                                SendKeys.SendWait(parsed[2]);
                            } else
                            if (parsed[1].Contains("Sleep"))
                            {
                                Thread.Sleep(Convert.ToInt32(parsed[2]));
                            }
                            Thread.Sleep(waitTime);
                        }
                    }
                }
                    
            }
        }

       

    }
}
