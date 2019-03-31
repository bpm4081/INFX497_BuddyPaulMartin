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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelDataReader;

namespace INFX497_BuddyPaulMartin
{
    public partial class Preview : Form
    {
        public Preview()
        {
            InitializeComponent();
        }

        public DataTable dataTable { get; set; }

        public MakeCalls makeCalls { get; set; }

        public bool calledFromMakeCalls = false;

        private void Preview_Load(object sender, EventArgs e)
        {
            dataGridView.DataSource = dataTable;
            if (calledFromMakeCalls == false)
            {
                makeCallsLabel.Visible = false;
            }
            else
            {
                makeCallsLabel.Text = "If you want to switch to a specific employee, then DOUBLE CLICK on the ROW HEADER column. It's the column with the black triangle.";
                makeCallsLabel.Location = new Point(25, 13);
                makeCallsLabel.Visible = true;
            }
        } 

        private void dataGridView_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            makeCallsLabel.Visible = false;
            loadingLabel.Visible = true;
            dataGridView.Visible = false;
            makeCalls.i = e.RowIndex - 1;
            makeCalls.insertEmployeeInfo();
            this.Close();
        }
    } //end of Preview Class
}