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
using System.Drawing;

namespace GuiGood
{
    public static class FunctionsLibrary
    {
        /// <summary>
       /// Add Colored Text to logs
       /// </summary>
       /// <param name="box"></param>
       /// <param name="text"></param>
       /// <param name="color"></param>
        public static void AppendText(RichTextBox box, string text, Color color)
        {

            int start = box.TextLength;
            box.AppendText(text);
            int length = text.Length;
            box.Select(start, length);
            box.SelectionColor = color;
        }

        /// <summary>
        /// Show Dialog Box and get input
        /// </summary>
        /// <param name="text"></param>
        /// <param name="caption"></param>
        /// <returns></returns>
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

    }
}
