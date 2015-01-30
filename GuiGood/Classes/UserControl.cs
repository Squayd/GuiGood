using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;

namespace GuiGood
{
    class UserControl
    {
        //Variables
        #region Variables
        private List<string> scripts = new List<string>();
        private int numberOfScripts;
        private Dictionary<string, Tuple<string, string>> elementList = new Dictionary<string, Tuple<string, string>>();
        private string scriptComment;
        private string scriptProcessName;
        private GuiGood gui;
        AutomationFunctions autoFunc = new AutomationFunctions();
        ListView listView2TEMP;
        private string ProcessNameTemp;
        #endregion


        /// <summary>
        /// Setup Application
        /// </summary>
        /// <param name="tabControl1"></param>
        /// <param name="listView1"></param>
        /// <param name="listView2"></param>
        /// <param name="treeView1"></param>
        /// <param name="imageList1"></param>
        /// <returns></returns>
        public bool setupApplication(TabControl tabControl1, ListView listView1, ListView listView2, TreeView treeView1, ImageList imageList1)
        {
            //Bring TabControl to Front
            tabControl1.BringToFront();

            //Load Default TreeView
            treeView1.ImageList = imageList1;

            //Add Columns
            listView1.View = View.Details;
            listView2.View = View.Details;
            int width = listView1.Width / 3;
            listView1.Columns.Add("Object", width, HorizontalAlignment.Left);
            listView1.Columns.Add("Invoke", width, HorizontalAlignment.Left);
            listView1.Columns.Add("Attributes", width, HorizontalAlignment.Left);
            listView2.Columns.Add("Objects", ((listView2.Width - 20) / 2), HorizontalAlignment.Left);
            listView2.Columns.Add("Type", ((listView2.Width - 20) / 2), HorizontalAlignment.Left);
            //Return
            listView2TEMP = listView2;
            return true;
        }

        /// <summary>
        /// Create A Project
        /// </summary>
        /// <param name="guiGoodLog"></param>
        /// <param name="treeView1"></param>
        /// <returns></returns>
        public Tuple<string,string> CreateProject(RichTextBox guiGoodLog, TreeView treeView1, string ProjectName)
        {
            string ProjectPath = "";
            //Create a new Project
            ProjectName = FunctionsLibrary.ShowDialog("Please enter the name of the Project:", "New Project");
            if (ProjectName == null || ProjectName == "")
            {
                return null;
            }
            if (Directory.Exists(ProjectName) == false && File.Exists(ProjectName) == false)
            {
                DirectoryInfo path = Directory.CreateDirectory(ProjectName);
                ProjectPath = Environment.CurrentDirectory.ToString() + "\\" + path.ToString();

            }
            else
            {
                FunctionsLibrary.AppendText(guiGoodLog, "That project name already exists. \n", Color.Red);
                MessageBox.Show("That name already exists.");
            }
            treeView1.Nodes.Clear();
            TreeNode rootNode = new TreeNode("Project Suite '" + ProjectName + "' (0 Script(s))");
            rootNode.ImageIndex = 0;
            treeView1.Nodes.Add(rootNode);
            treeView1.ExpandAll();
            FunctionsLibrary.AppendText(guiGoodLog, "Project created. \n", Color.Blue);
            return new Tuple<string, string>(ProjectName, ProjectPath);
        }

        /// <summary>
        /// Create a new Script
        /// </summary>
        /// <param name="ProjectPath"></param>
        /// <param name="guiGoodLog"></param>
        /// <param name="ProjectName"></param>
        /// <param name="scriptComboBox"></param>
        /// <param name="treeView1"></param>
        /// <returns></returns>
        public bool NewScript(string ProjectPath, RichTextBox guiGoodLog, string ProjectName, ComboBox scriptComboBox, TreeView treeView1)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.InitialDirectory = ProjectPath;
            saveDialog.Filter = "GuiScript (*.gs)|*.gs";
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                var myFile = File.Create(saveDialog.FileName);
                myFile.Dispose();
                loadScripts(ProjectPath, ProjectName, scriptComboBox, treeView1);
            }
            saveDialog.Dispose();
            FunctionsLibrary.AppendText(guiGoodLog, "Created a new script. \n", Color.Blue);
            return true;
        }

        /// <summary>
        /// Open a Project
        /// </summary>
        /// <param name="ProjectPath"></param>
        /// <param name="ProjectName"></param>
        /// <param name="scriptComboBox"></param>
        /// <param name="treeView1"></param>
        /// <param name="guiGoodLog"></param>
        /// <returns></returns>
        public Tuple<string,string> OpenProject(string ProjectPath, string ProjectName, ComboBox scriptComboBox, TreeView treeView1, RichTextBox guiGoodLog)
        {
            //Load Project
            FolderBrowserDialog openProject = new FolderBrowserDialog();
            openProject.SelectedPath = Environment.CurrentDirectory;
            if (openProject.ShowDialog() == DialogResult.OK)
            {
                ProjectPath = openProject.SelectedPath;
                ProjectName = System.IO.Path.GetFileName(openProject.SelectedPath);
                ScriptRunner.ProjectName = ProjectName;
                ScriptRunner.ProjectPath = ProjectPath;
                DirectoryInfo d = new DirectoryInfo(openProject.SelectedPath);
                scriptComboBox.Items.Clear();
                foreach (var file in d.GetFiles("*.gs"))
                {
                    scripts.Add(file.Name);
                    scriptComboBox.Items.Add(file.Name);
                }
            }

            numberOfScripts = scripts.Count();
            treeView1.Nodes.Clear();

            //Fill Tree
            TreeNode rootNode = new TreeNode("Project Suite '" + ProjectName + "' (" + numberOfScripts + " Script(s))");
            foreach (var script in scripts)
            {
                //Add Script Node
                TreeNode testNode = new TreeNode(script.ToString());
                testNode.ImageIndex = 1;
                testNode.SelectedImageIndex = 1;
                rootNode.Nodes.Add(testNode);
            }

            //root Node
            rootNode.ImageIndex = 0;
            treeView1.Nodes.Add(rootNode);
            treeView1.ExpandAll();
            openProject.Dispose();
            FunctionsLibrary.AppendText(guiGoodLog, "Loaded Project \n", Color.Blue);
            return new Tuple<string, string>(ProjectName, ProjectPath);
        }

        /// <summary>
        /// Save script 
        /// </summary>
        /// <param name="ProjectName"></param>
        /// <param name="ProjectPath"></param>
        /// <param name="listView1"></param>
        /// <param name="scriptComboBox"></param>
        /// <param name="tabControl1"></param>
        /// <param name="guiGoodLog"></param>
        /// <returns></returns>
        public bool SaveScript(string ProjectName, string ProjectPath, ListView listView1, ComboBox scriptComboBox, TabControl tabControl1, RichTextBox guiGoodLog)
        {
            //If Project File was Created!
            if (ProjectName != null)
            {
                ScriptRunner.ProjectName = ProjectName;
                ScriptRunner.ProjectPath = ProjectPath;
                switch (tabControl1.SelectedTab.Text)
                {
                    //If on Script Creator, save Script
                    case "Script Creator":
                        if (scriptComboBox.SelectedIndex != -1)
                        {
                            using (StreamWriter sw = new StreamWriter(ProjectPath + "\\" + scriptComboBox.SelectedItem.ToString(), false))
                            {
                                sw.WriteLine("//GuiScript -- Scripting Edition: 1.0 --");
                                int counter = 1;
                                foreach (ListViewItem item in listView1.Items)
                                {
                                    sw.WriteLine("//Path : " + counter.ToString());
                                    counter++;
                                    sw.WriteLine(item.SubItems[0].Text + "|" + item.SubItems[1].Text + "|" + item.SubItems[2].Text);
                                }
                                sw.WriteLine("//GuiScript -- END --");
                            }
                            FunctionsLibrary.AppendText(guiGoodLog, "Script saved. \n", Color.Blue);
                            return true;
                        }
                        else
                        {
                            FunctionsLibrary.AppendText(guiGoodLog, "No script selected. \n", Color.Red);
                            MessageBox.Show("No script selected.");
                            return false;
                        }
                        break;
                    //If on Project Viewer, save Project
                    case "Project Viewer":
                        return false;
                        break;
                }
                return false;
                
            }
            else
            {
                MessageBox.Show("You have not created a Project!");
                return false;
            }

        }

        /// <summary>
        /// Add functions to script Path
        /// </summary>
        /// <param name="comboBox3"></param>
        /// <param name="listView1"></param>
        /// <returns></returns>
        public bool AddFunctionsToPath(ComboBox comboBox3, ListView listView1)
        {
            if (comboBox3.SelectedIndex != -1)
            {
                if (comboBox3.SelectedItem.ToString() == "Write Memory Mapped File")
                {
                    string info = FunctionsLibrary.ShowDialog("Enter the name of the memory mapped file", "MMF");
                    string info2 = FunctionsLibrary.ShowDialog("Enter the information stored in the MMF:", "MMF");
                    ListViewItem item1 = new ListViewItem("Function");
                    item1.SubItems.Add(comboBox3.SelectedItem.ToString());
                    item1.SubItems.Add(info + ":" + info2);
                    listView1.Items.Add(item1);
                } else
                if (comboBox3.SelectedItem.ToString() == "Close Memory Mapped File")
                {
                    string info = FunctionsLibrary.ShowDialog("Enter the name of the memory mapped file", "MMF");
                    ListViewItem item1 = new ListViewItem("Function");
                    item1.SubItems.Add(comboBox3.SelectedItem.ToString());
                    item1.SubItems.Add(info);
                    listView1.Items.Add(item1);
                } else
                if (comboBox3.SelectedItem.ToString() == "Send Message")
                {
                    string info = FunctionsLibrary.ShowDialog("Enter the information stored in the variable:", "Variable Info");
                    ListViewItem item1 = new ListViewItem("Function");
                    item1.SubItems.Add(comboBox3.SelectedItem.ToString());
                    item1.SubItems.Add(info);
                    listView1.Items.Add(item1);
                } else
                if (comboBox3.SelectedItem.ToString() == "Sleep")
                {
                        string info = FunctionsLibrary.ShowDialog("Enter the time to sleep for:", "Sleep Info");
                        ListViewItem item1 = new ListViewItem("Function");
                        item1.SubItems.Add(comboBox3.SelectedItem.ToString());
                        item1.SubItems.Add(info);
                        listView1.Items.Add(item1);
                } else
                if (comboBox3.SelectedItem.ToString() == "END")
                {
                    ListViewItem item1 = new ListViewItem("Function");
                    item1.SubItems.Add(comboBox3.SelectedItem.ToString());
                    item1.SubItems.Add("END");
                    listView1.Items.Add(item1);
                }

            }
            return true;
        }

        /// <summary>
        /// View Script in Viewer
        /// </summary>
        /// <param name="e"></param>
        /// <param name="designView"></param>
        /// <param name="ProjectPath"></param>
        /// <returns></returns>
        public bool ViewScript(TreeNodeMouseClickEventArgs e, RichTextBox designView, string ProjectPath)
        {
            ScriptRunner.ProjectPath = ProjectPath;
            designView.Text = "";
            if (e.Node != null)
            {
                if (e.Node.Text.Contains(".gs"))
                {
                    ScriptRunner.ScriptName = e.Node.Text;
                    //Clicked Script
                    FunctionsLibrary.AppendText(designView, "GuiScript START :: \n", Color.Black);
                    FunctionsLibrary.AppendText(designView, "    [PATH] : \n", Color.Purple);
                    try
                    {
                        using (StreamReader sw = new StreamReader(ProjectPath + "\\" + e.Node.Text))
                        {
                            string line;
                            while ((line = sw.ReadLine()) != null)
                            {
                                if (!line.StartsWith("//"))
                                {

                                    //Parse
                                    string[] parsed = line.Split('|');
                                    //Object Name 
                                    FunctionsLibrary.AppendText(designView, "        ", Color.Black);
                                    FunctionsLibrary.AppendText(designView, " [NAME] ", Color.Blue);
                                    FunctionsLibrary.AppendText(designView, parsed[0], Color.Black);
                                    FunctionsLibrary.AppendText(designView, " [INVOKE] ", Color.Blue);
                                    FunctionsLibrary.AppendText(designView, parsed[1] + "\n", Color.Black);

                                }
                            }
                            FunctionsLibrary.AppendText(designView, "    [END PATH] : \n", Color.Purple);
                            FunctionsLibrary.AppendText(designView, "GuiScript END :: \n", Color.Black);
                        }
                        return true;
                    }
                    catch { return false; }

                }
                return false;
            }
            return false;
        }

        /// <summary>
        /// Load a script into Path
        /// </summary>
        /// <param name="scriptComboBox"></param>
        /// <param name="listView1"></param>
        /// <param name="ProjectPath"></param>
        /// <returns></returns>
        public bool LoadScriptIntoPath(ComboBox scriptComboBox, ListView listView1, string ProjectPath)
        {
            if (scriptComboBox.SelectedIndex != -1)
            {
                listView1.Items.Clear();
                ScriptRunner.ProjectPath = ProjectPath;
                using (StreamReader sw = new StreamReader(ProjectPath + "\\" + scriptComboBox.SelectedItem.ToString()))
                {
                    string line;
                    while ((line = sw.ReadLine()) != null)
                    {
                        if (!line.StartsWith("//"))
                        {

                            //Parse
                            string[] parsed = line.Split('|');
                            //Object Name 
                            ListViewItem item1 = new ListViewItem(parsed[0]);
                            //Object Invoke
                            item1.SubItems.Add(parsed[1]);
                            //Object Attributes
                            item1.SubItems.Add(parsed[2]);
                            listView1.Items.Add(item1);

                        }
                    }
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Change Invoke method of an Object in Path
        /// </summary>
        /// <param name="comboBox1"></param>
        /// <param name="listView1"></param>
        /// <returns></returns>
        public bool ChangeInvoke(ComboBox comboBox1, ListView listView1)
        {
            if (comboBox1.SelectedIndex != -1)
            {
                if (comboBox1.SelectedItem.ToString() != "Select Item" && comboBox1.SelectedItem.ToString() != "Set Text")
                {
                    foreach (ListViewItem item in listView1.SelectedItems)
                    {
                        if (comboBox1.SelectedItem.ToString() != null)
                        {
                            item.SubItems[0].Text = item.Text;
                            item.SubItems[1].Text = comboBox1.SelectedItem.ToString();
                        }
                    }
                }
                else
                {
                    //popup Box
                    string itemToSelect = FunctionsLibrary.ShowDialog("What is the text for the item?", "Item Text");
                    foreach (ListViewItem item in listView1.SelectedItems)
                    {
                        if (comboBox1.SelectedItem.ToString() != null)
                        {
                            item.SubItems[0].Text = item.Text;
                            item.SubItems[1].Text = comboBox1.SelectedItem.ToString() + ":" + itemToSelect;
                        }
                    }
                }
                return true;
            }
            return false;
        }

        /// <summary>
        /// View Attributes of an Object
        /// </summary>
        /// <param name="listView1"></param>
        /// <param name="objectName"></param>
        /// <param name="objectType"></param>
        /// <param name="objectID"></param>
        /// <param name="comboBox1"></param>
        /// <param name="comboBox2"></param>
        /// <param name="textBox1"></param>
        /// <param name="textBox2"></param>
        /// <returns></returns>
        public bool ViewAttributes(ListView listView1, Label objectName, Label objectType, Label objectID, ComboBox comboBox1, ComboBox comboBox2, TextBox textBox1, TextBox textBox2)
        {
            foreach (ListViewItem name in listView1.SelectedItems)
            {
                if (elementList.ContainsKey(name.Text))
                {
                    objectName.Text = name.Text;
                    objectType.Text = elementList[name.Text].Item1;
                    objectID.Text = elementList[name.Text].Item2;
                    comboBox1.SelectedIndex = comboBox1.FindStringExact(name.SubItems[1].Text);
                    //Parse Attributes
                    if (name.SubItems[2].Text != "null")
                    {
                        List<string> listString = ParseAttributes(name.SubItems[2].Text);
                        comboBox2.SelectedIndex = comboBox2.FindString(listString[0]);
                        textBox1.Text = scriptComment;
                        textBox2.Text = scriptProcessName;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// Load Objects into ListView
        /// </summary>
        public bool LoadObjects(string ProcessName, ListView listView2)
        {
            //Grab Process
            Process[] processes = Process.GetProcessesByName(ProcessName);
            //Foreach process
            foreach (Process p in processes)
            {
                //Declare Variables
                AutomationElement window = AutomationElement.FromHandle(p.MainWindowHandle);
                AutomationElementCollection listofelements = autoFunc.FindAllEnabled(window);

                foreach (AutomationElement elem in listofelements)
                {
                    if (elem.Current.Name != "")
                    {
                        AddListView(elem.Current.Name, elem.Current.LocalizedControlType.ToString(), listView2);
                    }
                }
                AddElementList(listofelements);
            }
            return true;
        }

        /// <summary>
        /// Scan Process for Objects
        /// </summary>
        /// <param name="ProcessName"></param>
        /// <param name="listView2"></param>
        /// <param name="userControl"></param>
        /// <param name="guiGoodLog"></param>
        /// <returns></returns>
        public bool ScanProcess(string ProcessName, ListView listView2, UserControl userControl, RichTextBox guiGoodLog)
        {
            listView2.Items.Clear();
            userControl.LoadObjects(ProcessName, listView2);
            ProcessNameTemp = ProcessName;
            if (listView2.Items.Count > 0)
            {
                FunctionsLibrary.AppendText(guiGoodLog, "Retrieved Components from Process: " + ProcessName + " ... " + listView2.Items.Count.ToString() + " Control(s) found. \n", Color.Blue);
            }
            return true;
        }

        /// <summary>
        /// Add List View
        /// </summary>
        /// <param name="name"></param>
        /// <param name="localtype"></param>
        public void AddListView(string name, string localtype, ListView listView2)
        {
            var item1 = new ListViewItem(new[] { name, localtype });
            if (listView2.InvokeRequired)
            {
                listView2.Invoke((Action)delegate { listView2.Items.Add(item1); });
            }
            else
            {
                listView2.Items.Add(item1);
            }
        }

        /// <summary>
        /// Set the GUI reference
        /// </summary>
        /// <param name="gui1"></param>
        public void SetGUI(GuiGood gui1)
        {
            gui = gui1;
        }

        /// <summary>
        /// Set Event Handlers
        /// </summary>
        public void SetEventHandler(string ProcessName, ListView listView2)
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
            LoadObjects(ProcessNameTemp, listView2TEMP);
            gui.SortColumn();
        }

        /// <summary>
        /// Set Form Objects to this Instance, (reference)
        /// </summary>
        /// <param name="listView2"></param>
        public void SetFormObjects(ListView listView2)
        {
            listView2TEMP = listView2;
        }

        /// <summary>
        /// Parse Attributes of text
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private List<string> ParseAttributes(string text)
        {
            List<string> result = new List<string>();
            string[] temp = text.Split('-');
            //Variable path
            string[] temp2 = temp[0].Split(':');
            string variablePath = temp2[1];
            result.Add(temp2[1]);
            //Comment
            string[] temp3 = temp[1].Split(':');
            string commentText = temp3[1];
            result.Add(temp3[1]);
            //Process
            string[] temp4 = temp[2].Split(':');
            result.Add(temp4[1]);

            //return Tuple
            return result;
        }

        /// <summary>
        /// Load Scripts Function
        /// </summary>
        /// <param name="path"></param>
        /// <param name="ProjectName"></param>
        /// <param name="scriptComboBox"></param>
        /// <param name="treeView1"></param>
        /// <returns></returns>
        private bool loadScripts(string path, string ProjectName, ComboBox scriptComboBox, TreeView treeView1)
        {
            DirectoryInfo d = new DirectoryInfo(path);
            scripts.Clear();
            scriptComboBox.Items.Clear();
            foreach (var file in d.GetFiles("*.GS"))
            {
                scripts.Add(file.Name);
                scriptComboBox.Items.Add(file.Name);
            }
            treeView1.Nodes.Clear();
            numberOfScripts = scripts.Count();
            //Fill Tree
            TreeNode rootNode = new TreeNode("Project Suite '" + ProjectName + "' (" + numberOfScripts + " Script(s))");
            foreach (var script in scripts)
            {
                //Add Script Node
                TreeNode testNode = new TreeNode(script.ToString());
                testNode.ImageIndex = 1;
                testNode.SelectedImageIndex = 1;
                rootNode.Nodes.Add(testNode);
            }

            //root Node
            rootNode.ImageIndex = 0;
            treeView1.Nodes.Add(rootNode);
            treeView1.ExpandAll();
            return true;

        }

        /// <summary>
        /// Sort elements in list
        /// </summary>
        /// <param name="collection"></param>
        public void AddElementList(AutomationElementCollection collection)
        {
            foreach (AutomationElement ele in collection)
            {
                if (ele.Current.Name != "")
                {
                    if (!elementList.ContainsKey(ele.Current.Name))
                    {
                        elementList.Add(ele.Current.Name, new Tuple<string, string>(ele.Current.LocalizedControlType.ToString(), ele.Current.AutomationId.ToString()));
                    }
                }
            }
        }


    }
}
