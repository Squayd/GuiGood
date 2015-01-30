﻿using System;
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
            return true;
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
