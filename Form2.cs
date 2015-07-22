using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetsuiteOnlineServicesOrders
{
    public partial class Form2 : Form
    {
        Hashtable htItems = new Hashtable();
        public Form2()
        {
            InitializeComponent();
        }

        public Form2(ArrayList arrItems)
        {
            InitializeComponent();

            htItems.Clear();
            listBox1.Items.Clear();
            foreach (string[] objItem in arrItems)
            {
                int numIndex = listBox1.Items.Add(objItem[1]);
                htItems.Add(numIndex, objItem[0]);
            }
        }

        public string ItemId
        {
            get
            {
                return htItems[listBox1.SelectedIndex].ToString();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex < 0)
            {
                Common.ShowMessage("Please, select a product!");
                button1.DialogResult = System.Windows.Forms.DialogResult.None;
                this.DialogResult = System.Windows.Forms.DialogResult.None;
            }
            else
            {
                Common.ShowMessage(htItems[listBox1.SelectedIndex].ToString());
                button1.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
            }
        }
    }
}
