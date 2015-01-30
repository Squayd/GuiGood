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

        //Variable
        #region Variables
        UserControl userControl = new UserControl();
        ListViewColumnSorter lvwColumnSorter = new ListViewColumnSorter();
        List<ListViewItem> items = new List<ListViewItem>();
        
        string ProjectName;
        string ProjectPath;
        int numberOfScripts;
        string scriptVariablePath;
        string scriptComment;
        string scriptProcessName;
        
        #endregion


        //Load Form
        #region Load
        //Load Form
        private void GuiGood_Load(object sender, EventArgs e)
        {
            //Initialize Instances
            toolStrip1.Renderer = new MySR();
            userControl.SetGUI(this);
            this.listView2.ListViewItemSorter = new ListViewItemComparer(1);
            

            //Setup Form
            userControl.setupApplication(tabControl1, listView1, listView2, treeView1, imageList1);
            
        }
        //Initialize Form
        public GuiGood()
        {
            InitializeComponent();
        }
        #endregion

        






























        //Control Events
        #region Control Events
        //Create a new Project
        private void newProjectToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Tuple<string,string> projectTuple = userControl.CreateProject(guiGoodLog, treeView1, ProjectName);
            if (projectTuple != null)
            {
                ProjectName = projectTuple.Item1;
                ProjectPath = projectTuple.Item2;
            }
        }

        //New Script
        private void newScriptToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            userControl.NewScript(ProjectPath, guiGoodLog, ProjectName, scriptComboBox, treeView1);
        }

        //Open a Project
        private void openProjectToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Tuple<string,string> tupleTemp = userControl.OpenProject(ProjectPath, ProjectName, scriptComboBox, treeView1, guiGoodLog);
            ProjectName = tupleTemp.Item1;
            ProjectPath = tupleTemp.Item2;
        }

        //Save Script
        private void saveToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            userControl.SaveScript(ProjectName, ProjectPath, listView1, scriptComboBox, tabControl1, guiGoodLog);
        }

        //Add function to path
        private void pictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            userControl.AddFunctionsToPath(comboBox3, listView1);
        }

        //View script in Viewer
        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            userControl.ViewScript(e, designView, ProjectPath);
        }

        //Load script into Path editor
        private void scriptComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            userControl.LoadScriptIntoPath(scriptComboBox, listView1, ProjectPath);
        }

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

        //Change Invoke of Object
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            userControl.ChangeInvoke(comboBox1, listView1);
        }

        //View Attributes of Object
        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            userControl.ViewAttributes(listView1, objectName, objectType, objectID, comboBox1, comboBox2, textBox1, textBox2);
        }

        //Item Drag
        private void listView2_ItemDrag(object sender, ItemDragEventArgs e)
        {
            DoDragDrop(e.Item, DragDropEffects.Move);
        }

        //Item Drop
        private void listView1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                var item = e.Data.GetData(typeof(ListViewItem)) as ListViewItem;
                var item1 = new ListViewItem(new[] { item.Text, "null", "null" });
                listView1.Items.Add(item1);
            }
        }

        //Item Drag Effect Enter
        private void listView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                e.Effect = DragDropEffects.Move;
            }
        }

        //Item Selected
        private void listView1_ItemActivate(object sender, EventArgs e)
        {
            var item = listView1.SelectedItems;
            foreach (object itemsel in item)
            {

                listView1.Items.Remove((ListViewItem)itemsel);
            }
        }

        //Sort Columns (Objects Scanned)
        private void listView2_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            this.listView2.ListViewItemSorter = new ListViewItemComparer(e.Column);
        }

        //Scan Process for Objects
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            userControl.ScanProcess(textBox3.Text, listView2, userControl, guiGoodLog);
            userControl.SetEventHandler(textBox3.Text, listView2);
        }

        #endregion


        //Column Functions
        #region ColumnFunctions
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
        #endregion

        private void toolStripSplitButton1_ButtonClick(object sender, EventArgs e)
        {
            ScriptSettings scriptSettings = new ScriptSettings();
            scriptSettings.Show();
        }

    }
}
