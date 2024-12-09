using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ExcelQueryAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // add RibbonDropDownItem to comboBox1
            RibbonDropDownItem item1 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            item1.Label = "Option 1 - This is a long option";
            comboBox1.Items.Add(item1);

            RibbonDropDownItem item2 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            item2.Label = "Option 2";
            comboBox1.Items.Add(item2);

            RibbonDropDownItem item3 = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            item3.Label = "Option 3";
            comboBox1.Items.Add(item3);
            //button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            //comboBox1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            
            
            


        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //MessageBox.Show("Button clicked.");
            // get the selected item from comboBox1

            //paramGroup.Items.Add(Globals.Factory.GetRibbonFactory().CreateRibbonEditBox());


            var selection = comboBox1.Text;
            MessageBox.Show(selection);
        }

        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Dialog Launcher clicked.");
        }
    }
}
