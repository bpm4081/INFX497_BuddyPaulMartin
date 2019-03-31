using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace INFX497_BuddyPaulMartin
{
    public partial class AddPayload : Form
    {
        public AddPayload()
        {
            InitializeComponent();
        }

        private void btnAddPayload_Click(object sender, EventArgs e)
        {
            try
            {
                if (!String.IsNullOrWhiteSpace(txtPayloadName.Text) && !String.IsNullOrWhiteSpace(rtxtAddPayload.Text))
                {
                    string[] lines = rtxtAddPayload.Lines;
                    string newpayload = Path.Combine(@"..\\..\\payloads\\", txtPayloadName.Text.ToString() + ".html");
                    using (StreamWriter newfile = new StreamWriter(newpayload))
                    {
                        foreach (string line in lines)
                        {
                            newfile.WriteLine(line.ToString());
                        }
                    }
                }
                lblResult.Visible = true;
                lblResult.ForeColor = System.Drawing.Color.Lime;
                lblResult.Text = "New payload added";
            }
            catch
            {
                lblResult.Visible = true;
                lblResult.ForeColor = System.Drawing.Color.Red;
                lblResult.Text = "Failed to add new payload";
            }
            
        }

        private void enableAddPayload()
        {
            if (rtxtAddPayload.Text != "" && txtPayloadName.Text != "")
            {
                lblResult.Visible = false;
                btnAddPayload.Enabled = true;
                btnAddPayload.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
                btnAddPayload.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
                btnAddPayload.ForeColor = System.Drawing.Color.White;
            } 
            else
            {
                lblResult.Visible = false;
                btnAddPayload.Enabled = false;
                btnAddPayload.BackColor = System.Drawing.Color.Gray;
                btnAddPayload.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
                btnAddPayload.ForeColor = System.Drawing.Color.LightGray;
            }
        }

        private void rtxtAddPayload_TextChanged(object sender, EventArgs e)
        {
            enableAddPayload();
        }

        private void txtPayloadName_TextChanged(object sender, EventArgs e)
        {
            enableAddPayload();
        }

        private void btnImportPayload_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog() { Filter = "HTML or Text File|*.html;*.txt", ValidateNames = true };
            DialogResult result = openFile.ShowDialog();
            //verify if file opened correctly
            if (result == DialogResult.OK)
            {
                string filePath = openFile.FileName;
                string[] splitter1 = filePath.Split('\\');
                string[] splitter2 = splitter1[splitter1.Length - 1].Split('.');
                string fileName = splitter2[0];
                //System.IO.File.Copy(Path.Combine(Path.GetDirectoryName(filePath)),  @"..\\..\\payloads\\" + fileName + ".html");
                System.IO.File.Copy(filePath, @"..\\..\\payloads\\" + fileName + ".html");
                lblResult.Visible = true;
                lblResult.ForeColor = System.Drawing.Color.Lime;
                lblResult.Text = "New payload added";
            }
            else
            {
                MessageBox.Show("Error copying payload.");
                lblResult.Visible = true;
                lblResult.ForeColor = System.Drawing.Color.Red;
                lblResult.Text = "Failed to add new payload";
            }
        }
    }
}