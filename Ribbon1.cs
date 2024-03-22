using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Latin_Cyrillic_TextConvert
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // no code
        }

        private void LatinToCyrillic_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.L_to_C_Text_Convert();
        }

        private void CyrillicToLatin_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.C_to_L_Text_Convert();
        }

        private void About_Click(object sender, RibbonControlEventArgs e)
        {
            Form aboutWindow = new Form();
            aboutWindow.Text = "About add-in";
            aboutWindow.AutoSize = true;
            aboutWindow.MaximizeBox = false;
            aboutWindow.MinimizeBox = false;
            Icon icon = Properties.Resources.info;
            aboutWindow.Icon = icon;
            string text = "This add-in was created as a graduation project. Its purpose is to help people " +
                "use and convert cyrillic and latin text as easy and fast as possible.\n\n" +
                "\nAuthor: Dragana Majstorović\nMentor: Saša Milić" +
                "\n\nSchool of electrical engineering Prijedor,\nJune 2021";
            Label labelText = new Label();
            labelText.Text = text;
            labelText.Padding = new Padding(20);
            labelText.MaximumSize = new Size(400,0);
            labelText.AutoSize = true;
            aboutWindow.Controls.Add(labelText);
            aboutWindow.Show();
        }
    }
}
