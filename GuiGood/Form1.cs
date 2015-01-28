using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;
using UIAutomationClient;

namespace GuiGood
{
    public partial class GuiGood : Form
    {

        //Variables
        #region Variables
        FunctionsLibrary funclib = new FunctionsLibrary();
        ListViewColumnSorter lvwColumnSorter = new ListViewColumnSorter();
        List<ListViewItem> items = new List<ListViewItem>();
        Dictionary<string, Tuple<string, string>> elementList = new Dictionary<string, Tuple<string, string>>();
        string ProjectName;
        string ProjectPath;
        int numberOfScripts;
        string scriptVariablePath;
        string scriptComment;
        string scriptProcessName;
        List<string> scripts = new List<string>();
        #endregion
        //Load Form
        #region Load
        //Load Stuffz
        private void GuiGood_Load(object sender, EventArgs e)
        {
            tabControl1.BringToFront();
            /*Cover2.Location = new Point(0, 673);
            Cover2.BringToFront();
            //this.Controls.SetChildIndex(Cover2, -100);
            Cover3.Location = new Point(0, 71);
            Cover3.BringToFront();
           // this.Controls.SetChildIndex(Cover3, -100); */
            
            //ToolStrip
            toolStrip1.Renderer = new MySR();
            //Load Default TreeView
            treeView1.ImageList = imageList1;
            

            //List Viewer
            //Add Columns
            listView1.View = View.Details;
            listView2.View = View.Details;
            int width = listView1.Width / 3;
            listView1.Columns.Add("Object", width, HorizontalAlignment.Left);
            listView1.Columns.Add("Invoke", width, HorizontalAlignment.Left);
            listView1.Columns.Add("Attributes", width, HorizontalAlignment.Left);
            listView2.Columns.Add("Objects", ((listView2.Width - 20) / 2), HorizontalAlignment.Left);
            listView2.Columns.Add("Type", ((listView2.Width - 20) / 2), HorizontalAlignment.Left);

            //Sort View
            this.listView2.ListViewItemSorter = new ListViewItemComparer(1);

            //Library
            funclib.SetGUI(this);
            funclib.LoadObjects();
            funclib.SetEventHandler();
            
           
        }
        public GuiGood()
        {
            InitializeComponent();
        }
        #endregion
        //Column Functions
        #region ColumnFunctions
        /// <summary>
        /// Sort Columns
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView2_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            this.listView2.ListViewItemSorter = new ListViewItemComparer(e.Column);
        }
        public void SortColumn()
        {
            if (listView2.InvokeRequired)
            {
                listView2.Invoke((Action)delegate { this.listView2.ListViewItemSorter = new ListViewItemComparer(1); });
            }
            else
            {
                this.listView2.ListViewItemSorter = new ListViewItemComparer(1);
            }
        }
        public void ClearColumn()
        {
            if (listView2.InvokeRequired)
            {
                listView2.Invoke((Action)delegate { listView2.Items.Clear(); });
            }
            else
            {
                //listView2.Items.Clear();
            }
        }
        /// <summary>
        /// Add List View Items
        /// </summary>
        /// <param name="name"></param>
        /// <param name="localtype"></param>
        public void AddListView(string name, string localtype)
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
        #endregion
        //Element Sorting
        #region Element Sorting
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
        #endregion
        //Drag Item
        #region Drag Item
        private void listView2_ItemDrag(object sender, ItemDragEventArgs e)
        {
            DoDragDrop(e.Item, DragDropEffects.Move);
        }

        private void listView1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                var item = e.Data.GetData(typeof(ListViewItem)) as ListViewItem;
                var item1 = new ListViewItem(new[] { item.Text, "null", "null" });
                listView1.Items.Add(item1);
            }
        }

        private void listView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                e.Effect = DragDropEffects.Move;
            }
        }

        private void listView1_ItemActivate(object sender, EventArgs e)
        {
            var item = listView1.SelectedItems;
            foreach (object itemsel in item)
            {

                listView1.Items.Remove((ListViewItem)itemsel);
            }
        }
        #endregion








        /// <summary>
        /// View Attributes
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
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
                        Tuple<string, string> tempTuple = ParseAttributes(name.SubItems[2].Text);
                        comboBox2.SelectedIndex = comboBox2.FindString(tempTuple.Item1);
                        textBox1.Text = scriptComment;
                        textBox2.Text = scriptProcessName;
                    }
                }
            }
        }
        /// <summary>
        /// Change Invoke Method on Object
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex != -1)
            {
                if (comboBox1.SelectedItem.ToString() != "Select Item")
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
                    string itemToSelect = ShowDialog("What is the Item Text you want to select?", "Item Selection");
                    foreach (ListViewItem item in listView1.SelectedItems)
                    {
                        if (comboBox1.SelectedItem.ToString() != null)
                        {
                            item.SubItems[0].Text = item.Text;
                            item.SubItems[1].Text = comboBox1.SelectedItem.ToString() + ":" + itemToSelect;
                        }
                    }
                }
            }
   
        }
        //Button Functions
        #region Button Functions
        //Save Button
        private void saveButton_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView1.SelectedItems)
            {
                item.SubItems[2].Text = "V:" + comboBox2.SelectedItem.ToString().Substring(0, 1) + "-C:" + textBox1.Text + "-P:" + textBox2.Text;
            }
        }
        //Clear Button
        private void clearButton_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            textBox1.Text = "";
            textBox2.Text = "";
        }
        #endregion

        //Parse Attributes
        private Tuple<string, string> ParseAttributes(string text)
        {
            string[] temp = text.Split('-');
            //Variable path
            string[] temp2 = temp[0].Split(':');
            string variablePath = temp2[1];
            scriptVariablePath = temp2[1];
            //Comment
            string[] temp3 = temp[1].Split(':');
            string commentText = temp3[1];
            scriptComment = temp3[1];
            //Process
            string[] temp4 = temp[2].Split(':');
            scriptProcessName = temp4[1];

            //return Tuple
            return new Tuple<string, string>(variablePath, commentText);
        }


        //Dialog Box
        #region Extra
        public static string ShowDialog(string text, string caption)
        {
            if (text.Length > 52)
            {
                string sub1 = text.Substring(0, 52) + "\n";
                string sub2 = text.Substring(53, text.Length - 53);
                text = sub1 + sub2;
            }
            Form prompt = new Form();
            prompt.Width = 500;
            prompt.Height = 150;
            prompt.FormBorderStyle = FormBorderStyle.FixedDialog;
            prompt.Text = caption;
            prompt.StartPosition = FormStartPosition.CenterScreen;
            Label textLabel = new Label() { Left = 50, Height = 200, Top = 5, AutoSize = true, Text = text };
            TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 400 };
            Button confirmation = new Button() { Text = "Ok", Left = 350, Width = 100, Top = 75 };
            confirmation.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;
            prompt.ShowDialog();
            return textBox.Text;
        }
        #endregion

   
        //Load Scripts Function
        private void loadScripts(string path)
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
            
        }
        //Load script into Path editor
        private void scriptComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (scriptComboBox.SelectedIndex != -1)
            {
                listView1.Items.Clear();
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
                }
            }
        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            designView.Text = "";
            if (e.Node != null)
            {
                if (e.Node.Text.Contains(".gs"))
                {
                        //Clicked Script
                        AppendText(designView, "GuiScript START :: \n", Color.Black);
                        AppendText(designView, "    [PATH] : \n", Color.Purple);
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
                                        AppendText(designView, "        ", Color.Black);
                                        AppendText(designView, " [NAME] ", Color.Blue);
                                        AppendText(designView, parsed[0], Color.Black);
                                        AppendText(designView, " [INVOKE] ", Color.Blue);
                                        AppendText(designView, parsed[1] + "\n", Color.Black);

                                    }
                                }
                                AppendText(designView, "    [END PATH] : \n", Color.Purple);
                                AppendText(designView, "GuiScript END :: \n", Color.Black);
                            }
                        }
                        catch { }

                }
            }
        }
        public static void AppendText(RichTextBox box, string text, Color color)
        {
            
            int start = box.TextLength;
            box.AppendText(text);
            int length = text.Length;
            box.Select(start, length);
            box.SelectionColor = color; 
        }

        private void newProjectToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            //Create a new Project
            ProjectName = ShowDialog("Please enter the name of the Project:", "New Project");

            if (Directory.Exists(ProjectName) == false && File.Exists(ProjectName) == false)
            {
                DirectoryInfo path = Directory.CreateDirectory(ProjectName);
                ProjectPath = Environment.CurrentDirectory.ToString() + "\\" + path.ToString();

            }
            else
            {
                AppendText(guiGoodLog, "That project name already exists. \n", Color.Red);
                MessageBox.Show("That name already exists.");
            }
            treeView1.Nodes.Clear();
            TreeNode rootNode = new TreeNode("Project Suite '" + ProjectName + "' (0 Script(s))");
            rootNode.ImageIndex = 0;
            treeView1.Nodes.Add(rootNode);
            treeView1.ExpandAll();
            AppendText(guiGoodLog, "Project created. \n", Color.Blue);
        }

        private void newScriptToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.InitialDirectory = ProjectPath;
            saveDialog.Filter = "GuiScript (*.gs)|*.gs";
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                File.Create(saveDialog.FileName);
                loadScripts(ProjectPath);
            }
            saveDialog.Dispose();
            AppendText(guiGoodLog, "Created a new script. \n", Color.Blue);
            
        }

        private void openProjectToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            //Load Project
            FolderBrowserDialog openProject = new FolderBrowserDialog();
            openProject.SelectedPath = Environment.CurrentDirectory;
            if (openProject.ShowDialog() == DialogResult.OK)
            {
                ProjectPath = openProject.SelectedPath;
                ProjectName = System.IO.Path.GetFileName(openProject.SelectedPath);
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
            AppendText(guiGoodLog, "Loaded Project \n", Color.Blue);
        }

        private void saveToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            //If Project File was Created!
            if (ProjectName != null)
            {
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
                            AppendText(guiGoodLog, "Script saved. \n", Color.Blue);
                        }
                        else
                        {
                            AppendText(guiGoodLog, "No script selected. \n", Color.Red);
                            MessageBox.Show("No script selected.");
                        }
                        break;
                    //If on Project Viewer, save Project
                    case "Project Viewer":

                        break;
                }
            }
            else
            {
                MessageBox.Show("You have not created a Project!");
            }

        }

        private void GuiGood_SizeChanged(object sender, EventArgs e)
        {
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            funclib.ProcessName = textBox3.Text;
            listView2.Items.Clear();
            funclib.LoadObjects();
            if (listView2.Items.Count > 0)
            {
                AppendText(guiGoodLog, "Retrieved Components from Process: " + textBox3.Text + " ... " + listView2.Items.Count.ToString() + " Control(s) found. \n", Color.Blue);  
            }
        }



    }
    #region ClassComparer
    /// <summary>
    /// Comparererererer for listview
    /// </summary>
    class ListViewItemComparer : IComparer
    {
        private int col = 0;

        public ListViewItemComparer(int column)
        {
            col = column;
        }
        public int Compare(object x, object y)
        {
            return String.Compare(((ListViewItem)x).SubItems[col].Text, ((ListViewItem)y).SubItems[col].Text);
        }
    }
#endregion
    public class MySR : ToolStripSystemRenderer
    {
        public MySR() { }

        protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
        {
            //base.OnRenderToolStripBorder(e);
        }
    }
}
