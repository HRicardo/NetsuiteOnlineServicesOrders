using System;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetsuiteOnlineServicesOrders
{
    public partial class Form1 : Form
    {
        delegate void SetTextCallback();
        delegate void SetText2Callback();

        public Form1()
        {
            InitializeComponent();
            Common.objForm = this;
        }

        private void cargarArchivoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                Common.dsData = new DataSet();
                Common.LoadExcelData(openFileDialog1.FileName);
                string[] arrCustomers = GlobalSettings.Default.exSheet.Split(new string[] { "," }, StringSplitOptions.None);
                comboBox1.Items.Clear();
                foreach (string strCustom in arrCustomers)
                {
                    comboBox1.Items.Add(strCustom);
                }
                button1.Enabled = true;
                try
                {
                    File.Copy(openFileDialog1.FileName, GlobalSettings.Default.pathLog + @"\load_" + DateTime.Now.Ticks + "_" + openFileDialog1.SafeFileName);
                }
                catch (Exception)
                {
                }
                Common.ShowMessage("File is loaded!");
                Cursor.Current = Cursors.Default;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            dataGridView1.Size = new Size(this.Size.Width - 50, this.Size.Height - 120);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.DataSource = Common.dsData.Tables[comboBox1.SelectedIndex];
                dataGridView1.ScrollBars = ScrollBars.Both;
            }
            catch (Exception)
            {
                Common.ShowMessage("The tab is empty.");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex > -1)
            {
                DataTable dtData = Common.dsData.Tables[comboBox1.SelectedIndex];
                progressBar1.Minimum = 0;
                progressBar1.Value = 0;
                progressBar1.Maximum = dtData.Rows.Count;
                this.button1.Enabled = false;
                this.comboBox1.Enabled = false;
                if (comboBox1.SelectedItem.Equals("format1800"))
                {
                    Common.LoadTable1800(dtData);
                }
                else
                {
                    if (comboBox1.SelectedItem.Equals("FTD"))
                        Common.LoadTableMassMarket(dtData);
                    else
                        Common.LoadTable(dtData);
                }
            }
        }

        public void IncrementProcess()
        {
            if (this.progressBar1.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(IncrementProcess);
                this.Invoke(d, new object[] { });
            }
            else
            {
                dataGridView1.ClearSelection();
                int numIndex = progressBar1.Value - 1;
                if (this.progressBar1.Value < this.progressBar1.Maximum)
                {
                    dataGridView1.Rows[progressBar1.Value].Selected = true;
                    this.progressBar1.Value += 1;
                }
                else
                {
                    numIndex = progressBar1.Value - 1;
                    this.button1.Enabled = true;
                    this.comboBox1.Enabled = true;               
                }
                this.lblTotal.Text = progressBar1.Value + " / " + progressBar1.Maximum;
                try
                {
                    dataGridView1.Rows[numIndex].ErrorText = Common.strMessage;
                    switch (Common.LastProcessResult)
                    {
                        case Common.ProcessResult.Processed:
                            dataGridView1.Rows[numIndex].DefaultCellStyle.BackColor = Color.Silver;
                            break;
                        case Common.ProcessResult.Ok1:
                            dataGridView1.Rows[numIndex].DefaultCellStyle.BackColor = Color.Yellow;
                            break;
                        case Common.ProcessResult.Ok2:
                            dataGridView1.Rows[numIndex].DefaultCellStyle.BackColor = Color.LightGreen;
                            break;
                        case Common.ProcessResult.Added:
                            dataGridView1.Rows[numIndex].DefaultCellStyle.BackColor = Color.LightBlue;
                            break;
                        case Common.ProcessResult.Error:
                            dataGridView1.Rows[numIndex].DefaultCellStyle.BackColor = Color.Orange;
                            break;
                        case Common.ProcessResult.Closed:
                            dataGridView1.Rows[numIndex].DefaultCellStyle.BackColor = Color.LightSalmon;
                            break;
                    }
                    //dataGridView1.Rows[numIndex].Selected = true;
                }
                catch (Exception)
                {
                }
            }
        }

        private void testLoginToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string strResult = Util.Login(GlobalSettings.Default.nsUser, GlobalSettings.Default.nsPassword, GlobalSettings.Default.nsAccount, GlobalSettings.Default.nsRole);
            //Util.GetPackedBunch("QB00003109");
            if (!string.IsNullOrEmpty(strResult))
            {
                Common.ShowMessage(GlobalSettings.Default.nsUser + ": " + strResult + "\n" + Util._service.Url);
            }
            else
            {
                Common.ShowMessage(GlobalSettings.Default.nsUser + ": Login OK!\n" + Util._service.Url);
            }
        }

        private void loadMassMarketFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                Common.dsData = new DataSet();
                Common.LoadExcelData(openFileDialog1.FileName);
                string[] arrCustomers = GlobalSettings.Default.exSheet.Split(new string[] { "," }, StringSplitOptions.None);
                comboBox1.Items.Clear();
                foreach (string strCustom in arrCustomers)
                {
                    comboBox1.Items.Add(strCustom);
                }
                button1.Enabled = true;
                try
                {
                    File.Copy(openFileDialog1.FileName, GlobalSettings.Default.pathLog + @"\load_" + DateTime.Now.Ticks + "_" + openFileDialog1.SafeFileName);
                }
                catch (Exception)
                {
                }
                Common.ShowMessage("File is loaded!");
                Cursor.Current = Cursors.Default;
                
            }

        }

        private void load1800FileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;
                button1.Enabled = true;
                string strNewFile = GlobalSettings.Default.pathLog + @"\load_1800_" + DateTime.Now.Ticks + ".csv";
                try                    
                {
                    //File.Copy(openFileDialog1.FileName, GlobalSettings.Default.pathLog + @"\load_1800_" + DateTime.Now.Ticks + "_" + openFileDialog1.SafeFileName);
                    File.Copy(openFileDialog1.FileName, strNewFile);
                }
                catch (Exception)
                {
                }

                Common.dsData = new DataSet();
                Common.LoadCSVData(strNewFile);
                comboBox1.Items.Clear();
                comboBox1.Items.Add("format1800");
                Common.ShowMessage("File is loaded!");
                Cursor.Current = Cursors.Default;
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }


    }
}
