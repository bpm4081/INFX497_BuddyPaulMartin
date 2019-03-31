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
using ExcelDataReader;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Threading;

namespace INFX497_BuddyPaulMartin
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        public DataTable dataTable;
        StreamReader streamer;
        string contactpath = string.Empty;
        string reportpath = string.Empty;
        string path = string.Empty;
        string itemText = string.Empty;
        string fileType = string.Empty;
        int currentYear = DateTime.Now.Year;
        Excel.Application xlApp;
        Word.Application wdApp;

        private void Main_Load(object sender, EventArgs e)
        {

        }

        // **METHOD THAT OPENS FILE EXPLORER AND FOCUSES THE NEWLY SAVED ITEM BY THE USER
        private void OpenFolder(string folderPath)
        {
            if (File.Exists(folderPath))
            {
                Process.Start(new ProcessStartInfo("explorer.exe", " /select, " + folderPath));
            }
        }

        // ####################################################################################################################################################################################################
        // ####################################################################################################################################################################################################
        // ###########################                                                                                                                                              ###########################
        // ###########################                                                            Comma Seperated File                                                              ###########################
        // ###########################                                                                                                                                              ###########################
        // ####################################################################################################################################################################################################
        // ####################################################################################################################################################################################################

        //Open Word Document File 
        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            //Tooltip for button for explnation
            ToolTip toolOpenFile = new ToolTip();
            toolOpenFile.ShowAlways = true;
            toolOpenFile.SetToolTip(btnOpenFile, "Open .docx Scoping Form");
            try
            {
                //open file dialog to select .docx Scoping Form
                OpenFileDialog openFile = new OpenFileDialog() { Filter = "Word Document|*.doc;*.docx", ValidateNames = true };
                //verify file opened 
                DialogResult result = openFile.ShowDialog();
                if (result == DialogResult.OK)
                {
                    //set fileType as Word and then create the Data Table that will be displayed in the preview window and store the Data Table into the class variable dataTable
                    fileType = "Word";
                    contactpath = openFile.FileName;
                    dataTable = wordDocToDataTable(contactpath);

                    //if file opens correctly then enable buttons to create CSV file or Preview data
                    btnPreview.Visible = true;
                    btnPreview2.Visible = false;
                    btnInsightCSV.Enabled = true;
                    btnKnowbe4CSV.Enabled = true;
                    txtUserGroup.Enabled = true;
                    btnCreateCallList.Enabled = true;
                    //change colors to show enabled buttons
                    txtUserGroup.BackColor = System.Drawing.Color.White;
                    btnInsightCSV.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
                    btnInsightCSV.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
                    btnInsightCSV.ForeColor = System.Drawing.Color.White;
                    btnKnowbe4CSV.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
                    btnKnowbe4CSV.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
                    btnKnowbe4CSV.ForeColor = System.Drawing.Color.White;
                    btnCreateCallList.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
                    btnCreateCallList.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
                    btnCreateCallList.ForeColor = System.Drawing.Color.White;
                    //move label to show successful file open
                    lblPullContacts.Left = 3;
                    lblPullContacts.Top = 70;
                    lblPullContacts.Visible = true;
                    lblPullContacts.ForeColor = System.Drawing.Color.Lime;
                    lblPullContacts.Text = "Success extracting data";
                }

            }
            //if opening file fails then disable buttons 
            catch
            {
                //hidden/disable buttons
                btnPreview.Visible = false;
                btnPreview2.Visible = false;
                btnInsightCSV.Enabled = false;
                btnKnowbe4CSV.Enabled = false;
                txtUserGroup.Enabled = false;
                btnCreateCallList.Enabled = false;
                //change colors to look disabled
                txtUserGroup.BackColor = System.Drawing.Color.Gray;
                btnInsightCSV.BackColor = System.Drawing.Color.Gray;
                btnInsightCSV.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
                btnInsightCSV.ForeColor = System.Drawing.Color.LightGray;
                btnKnowbe4CSV.BackColor = System.Drawing.Color.Gray;
                btnKnowbe4CSV.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
                btnKnowbe4CSV.ForeColor = System.Drawing.Color.LightGray;
                btnCreateCallList.BackColor = System.Drawing.Color.Gray;
                btnCreateCallList.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
                btnCreateCallList.ForeColor = System.Drawing.Color.LightGray;
                //notify user the file failed to extract
                lblPullContacts.Left = 3;
                lblPullContacts.Top = 70;
                lblPullContacts.Visible = true;
                lblPullContacts.ForeColor = System.Drawing.Color.Red;
                lblPullContacts.Text = "Failed extracting data";
                MessageBox.Show("The file could not be loaded");
            }
        }

        //import excel workbooks to get employee contact info 
        private void btnOpenExcelFile_Click(object sender, EventArgs e)
        {
            ToolTip toolOpenFile = new ToolTip();
            toolOpenFile.ShowAlways = true;
            toolOpenFile.SetToolTip(btnOpenExcelFile, "Open .xlsx Scoping Form");
            try
            {
                //open file dialog to select Excel scoping form
                OpenFileDialog openFile = new OpenFileDialog() { Filter = "Excel Workbook|*.xls;*.xlsx;*.csv", ValidateNames = true };
                DialogResult result = openFile.ShowDialog();
                //verify if file opened correctly
                if (result == DialogResult.OK)
                {
                    //set fileType as Excel and then create the Data Table that will be displayed in the preview window and store the Data Table into the class variable dataTable
                    fileType = "Excel";
                    contactpath = openFile.FileName;
                    dataTable = excelSheetToDataTable(contactpath, true);

                    //enable buttons 
                    btnPreview2.Visible = true;
                    btnPreview.Visible = false;
                    btnInsightCSV.Enabled = true;
                    btnKnowbe4CSV.Enabled = true;
                    txtUserGroup.Enabled = true;
                    btnCreateCallList.Enabled = true;
                    //change button colors to look enabled 
                    txtUserGroup.BackColor = System.Drawing.Color.White;
                    btnInsightCSV.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
                    btnInsightCSV.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
                    btnInsightCSV.ForeColor = System.Drawing.Color.White;
                    btnKnowbe4CSV.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
                    btnKnowbe4CSV.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
                    btnKnowbe4CSV.ForeColor = System.Drawing.Color.White;
                    btnCreateCallList.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
                    btnCreateCallList.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
                    btnCreateCallList.ForeColor = System.Drawing.Color.White;
                    //move open file label to notify user file successfuly opened
                    lblPullContacts.Left = 213;
                    lblPullContacts.Top = 68;
                    lblPullContacts.Visible = true;
                    lblPullContacts.ForeColor = System.Drawing.Color.Lime;
                    lblPullContacts.Text = "Success extracting data";
                }

            }
            //if file does not open correctly
            catch
            {
                //hide/disable buttons
                btnPreview2.Visible = false;
                btnPreview.Visible = false;
                btnInsightCSV.Enabled = false;
                btnKnowbe4CSV.Enabled = false;
                txtUserGroup.Enabled = false;
                btnCreateCallList.Enabled = false;
                //change button colors to look disabled
                txtUserGroup.BackColor = System.Drawing.Color.Gray;
                btnInsightCSV.BackColor = System.Drawing.Color.Gray;
                btnInsightCSV.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
                btnInsightCSV.ForeColor = System.Drawing.Color.LightGray;
                btnKnowbe4CSV.BackColor = System.Drawing.Color.Gray;
                btnKnowbe4CSV.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
                btnKnowbe4CSV.ForeColor = System.Drawing.Color.LightGray;
                btnCreateCallList.BackColor = System.Drawing.Color.Gray;
                btnCreateCallList.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
                btnCreateCallList.ForeColor = System.Drawing.Color.LightGray;
                //move open file label to notify user the file failed to open
                lblPullContacts.Left = 213;
                lblPullContacts.Top = 68;
                lblPullContacts.Visible = true;
                lblPullContacts.ForeColor = System.Drawing.Color.Red;
                lblPullContacts.Text = "Failed extracting data";
                MessageBox.Show("The file could not be loaded");
            }
        }

        //converts a Word Document Scoping form into a Data Table
        private DataTable wordDocToDataTable(string filePath)
        {
            Read displayData = new Read();
            displayData.NameFile = filePath;
            //get headers
            string[] copyHeader = displayData.WordTableHeader();
            //get table data 
            string[,] displayArray = displayData.WordDoc();

            //see if First Name and Last Name are in seperate columns
            int firstNameColumn = -1;
            int lastNameColumn = -1;
            string firstNameColumnName = null;
            string lastNameColumnName = null;
            for (int i = 0; i < copyHeader.Length; i++)
            {
                if (copyHeader[i].ToLower().Contains("first"))
                {

                    firstNameColumn = i;
                    firstNameColumnName = copyHeader[i];
                }
                else if (copyHeader[i].ToLower().Contains("last"))
                {
                    lastNameColumn = i;
                    lastNameColumnName = copyHeader[i];
                }
            }


            DataTable result = new DataTable();
            //create DataTable
            if (firstNameColumn != -1 & lastNameColumn != -1)
            { //create DataTable if First Name and Last Name are in SEPERATE columns
                //add headers to the columns in the DataTable
                result.Columns.Add("Name", typeof(String));
                for (int i = 0; i < copyHeader.Length; i++)
                {
                    if (!copyHeader[i].Equals(firstNameColumnName) & !copyHeader[i].Equals(lastNameColumnName))
                    {
                        result.Columns.Add(copyHeader[i], typeof(String));
                    }
                }
                //add employee info to the DataTable
                for (int i = 0; i < (displayArray.Length / copyHeader.Length); i++)
                {
                    DataRow row = result.NewRow();
                    row["Name"] = displayArray[i, firstNameColumn] + " " + displayArray[i, lastNameColumn];
                    for (int j = 0; j < copyHeader.Length; j++)
                    {
                        if (!copyHeader[j].Equals(firstNameColumnName) & !copyHeader[j].Equals(lastNameColumnName))
                        {
                            row[copyHeader[j]] = displayArray[i, j];
                        }
                    }
                    result.Rows.Add(row);
                }

            }
            else
            { //create DataTable if First Name and Last Name are in the SAME column
                //add headers to the columns in the DataTable
                for (int i = 0; i < copyHeader.Length; i++)
                {
                    result.Columns.Add(copyHeader[i], typeof(String));
                }
                //add employee info to the DataTable
                for (int i = 0; i < (displayArray.Length / copyHeader.Length); i++)
                {
                    DataRow row = result.NewRow();
                    for (int j = 0; j < copyHeader.Length; j++)
                    {
                        row[copyHeader[j]] = displayArray[i, j];
                    }
                    result.Rows.Add(row);
                }
            }
            return result;
        }

        //converts a Excel Workbook Scoping form into a Data Table
        private DataTable excelSheetToDataTable(string filePath, bool useFirstRowAsHeaders)
        {
            var file = new FileInfo(filePath);
            IExcelDataReader reader;
            FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.Read);
            if (file.Extension.Equals(".xls"))
                reader = ExcelReaderFactory.CreateBinaryReader(fs);
            else if (file.Extension.Equals(".xlsx"))
                reader = ExcelReaderFactory.CreateOpenXmlReader(fs);
            else if (file.Extension.Equals(".csv"))
                reader = ExcelReaderFactory.CreateCsvReader(fs);
            else
                throw new Exception("Invalid FileName");

            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = useFirstRowAsHeaders
                }
            };

            var dataSet = reader.AsDataSet(conf);
            var dt = dataSet.Tables[0];
            reader.Close();

            List<string> initialHeaders = new List<string>();
            foreach (DataColumn column in dt.Columns)
            {
                initialHeaders.Add(column.ColumnName);
            }
            int firstNameColumn = -1;
            int lastNameColumn = -1;
            string firstNameColumnName = null;
            string lastNameColumnName = null;
            for (int i = 0; i < initialHeaders.Count; i++)
            {
                if (initialHeaders[i].ToLower().Contains("first"))
                {
                    firstNameColumn = i;
                    firstNameColumnName = initialHeaders[i];
                }
                else if (initialHeaders[i].ToLower().Contains("last"))
                {
                    lastNameColumn = i;
                    lastNameColumnName = initialHeaders[i];
                }
            }

            DataTable result = new DataTable();
            List<String> headers = new List<String>();
            if (firstNameColumn != -1 & lastNameColumn != -1)
            { //create DataTable if First Name and Last Name are in SEPERATE columns
                result.Columns.Add("Name", typeof(String));
                foreach (DataColumn column in dt.Columns)
                {
                    if (!column.ColumnName.Equals(firstNameColumnName) & !column.ColumnName.Equals(lastNameColumnName))
                    {
                        result.Columns.Add(column.ColumnName, typeof(String));
                    }
                }
                foreach (DataRow dr in dt.Rows)
                {
                    DataRow row = result.NewRow();
                    row["Name"] = dr[firstNameColumnName].ToString() + " " + dr[lastNameColumnName];
                    foreach (DataColumn column in dt.Columns)
                    {
                        if (!column.ColumnName.Equals(firstNameColumnName) & !column.ColumnName.Equals(lastNameColumnName))
                        {
                            row[column.ColumnName] = dr[column.ColumnName];
                        }
                    }
                    result.Rows.Add(row);
                }
                foreach (DataColumn column in result.Columns)
                {
                    headers.Add(column.ColumnName);
                }
            }
            else
            { //create DataTable if First Name and Last Name are in the SAME column
                result = dt;
                headers = initialHeaders;
            }
            fs.Close();
            return result;
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            ToolTip toolPreview = new ToolTip();
            toolPreview.ShowAlways = false;
            toolPreview.SetToolTip(btnPreview, "Preview Extracted Data");
            Preview newpreview = new Preview();
            newpreview.dataTable = dataTable;
            newpreview.ShowDialog();
        }

        private void btnPreview2_Click(object sender, EventArgs e)
        {
            ToolTip toolPreview = new ToolTip();
            toolPreview.ShowAlways = false;
            toolPreview.SetToolTip(btnPreview, "Preview Extracted Data");
            Preview newpreview = new Preview();
            newpreview.dataTable = dataTable;
            newpreview.ShowDialog();
        }
        // ####################################################################################################################################################################################################
        // ####################################################################################################################################################################################################
        // ###########################                                                                                                                                              ###########################
        // ###########################                                                              CREATE CSV FILE                                                                 ###########################
        // ###########################                                                                                                                                              ###########################
        // ####################################################################################################################################################################################################
        // ####################################################################################################################################################################################################

        private void btnInsightCSV_Click(object sender, EventArgs e)
        {
            string userGroup = txtUserGroup.Text.ToString();
            Read reading = new Read();

            // WORD VALUES            
            if (fileType == "Word")
            {

                reading.NameFile = contactpath;
                // CREATE AN ARRAY TO HOLD HEADER VALUES FROM FILE
                string[] copyHeader = reading.WordTableHeader();
                // CREATE AN ARRAY TO HOLD DATA VALUES FROM FILE
                string[,] copyData = reading.WordDoc();

                // CREATE INDEXES FOR # OF ROWS AND # OF COLUMNS
                int rowcount = copyData.GetUpperBound(0) + 1;
                int colcount = copyData.GetUpperBound(1) + 1;
                // CREATE AN ARRAY TO PLACE DATA VALUES IN DESIRED ORDER
                string[,] reorderData = new string[rowcount, 7];
                try
                {
                    for (int i = 0; i < rowcount; i++)
                    {
                        for (int j = 0; j < colcount; j++)
                        {
                            // BLANK VALUES
                            
                            // SET MIDDLE NAME VALUE
                            reorderData[i, 1] = " ";
                            // SET USERGROUP VALUE
                            if (!String.IsNullOrWhiteSpace(userGroup))
                            {
                                reorderData[i, 6] = userGroup;
                            }
                            else
                            {
                                reorderData[i, 6] = " ";
                            }
                            // SET FIRST NAME VALUE
                            if (copyHeader[j].Contains("first"))
                            {
                                reorderData[i, 0] = copyData[i, j];
                            }
                            // SET LAST NAME VALUE
                            else if (copyHeader[j].Contains("last"))
                            {
                                reorderData[i, 2] = copyData[i, j];
                            }
                            // SET TITLE VALUE
                            else if (copyHeader[j].Contains("title"))
                            {
                                reorderData[i, 3] = copyData[i, j];
                            }
                            // SET PHONE NUMBER VALUE
                            else if (copyHeader[j].Contains("phone"))
                            {
                                reorderData[i, 4] = copyData[i, j];
                            }
                            // SET EMAIL ADDRESS VALUE
                            else if (copyHeader[j].Contains("email") | copyHeader[j].Contains("e-mail"))
                            {
                                reorderData[i, 5] = copyData[i, j];
                            }
                            else
                            {

                            }
                        }
                    }

                    string thisfile = String.Empty;
                    // CREATE A FILE SAVE DIALOG WITH DESIRED FILE FORMAT AND EXTENSION
                    SaveFileDialog fileStream = new SaveFileDialog();
                    fileStream.FileName = "insightupload.csv";
                    fileStream.DefaultExt = ".csv";
                    fileStream.Filter = "Comma Separated files (*.csv)|*.csv";
                    // DISPLAY THE CREATE FILE SAVE DIALOG BOX TO THE USER
                    DialogResult result = fileStream.ShowDialog();
                    // OBTAIN THE SAVE FILE NAME/LOCATION FROM USER INPUT  
                    if (result == DialogResult.OK)
                    {
                        thisfile = fileStream.FileName;
                    }
                    // CALL CREATE CLASS AND ASSIGN VALUES FOR READ FILE, SAVE FILE, AND PROPERLY-ORDERED DATA
                    Create makeFile = new Create(fileType, contactpath, thisfile, reorderData);
                    // CALL CREATE CLASS'S CSV-MAKING METHOD
                    makeFile.InsightUpload();
                    OpenFolder(thisfile);
                }
                catch
                {
                    MessageBox.Show("Unable to export data to file properly.");
                }
            }
            
            // EXCEL VALUES
            else if (fileType == "Excel")
            {
                reading.NameFile = contactpath;
                // CREATE AN ARRAY TO HOLD HEADER VALUES FROM FILE
                string[] copyHeader = reading.ExcelTableHeader();
                // CREATE AN ARRAY TO HOLD DATA VALUES FROM FILE
                string[,] copyData = reading.ExcelDoc();
                // CREATE INDEXES FOR # OF ROWS AND # OF COLUMNS
                int rowcount = copyData.GetUpperBound(0);
                int colcount = copyData.GetUpperBound(1) + 1;
                // CREATE AN ARRAY TO PLACE DATA VALUES IN DESIRED ORDER
                string[,] reorderData = new string[rowcount, 7];

                for (int i = 0; i < rowcount; i++)
                {
                    for (int j = 0; j < colcount; j++)
                    {
                        // BLANK VALUES

                        // SET MIDDLE NAME VALUE
                        reorderData[i, 1] = " ";
                        // SET USERGROUP VALUE
                        if (!String.IsNullOrWhiteSpace(userGroup))
                        {
                            reorderData[i, 6] = userGroup;
                        }
                        else
                        {
                            reorderData[i, 6] = " ";
                        }
                        // SET FIRST NAME VALUE
                        if (copyHeader[j].Contains("first"))
                        {
                            reorderData[i, 0] = copyData[i, j];
                        }
                        // SET LAST NAME VALUE
                        else if (copyHeader[j].Contains("last"))
                        {
                            reorderData[i, 2] = copyData[i, j];
                        }
                        // SET TITLE VALUE
                        else if (copyHeader[j].Contains("title"))
                        {
                            reorderData[i, 3] = copyData[i, j];
                        }
                        // SET PHONE NUMBER VALUE
                        else if (copyHeader[j].Contains("phone"))
                        {
                            reorderData[i, 4] = copyData[i, j];
                        }
                        // SET EMAIL ADDRESS VALUE
                        else if (copyHeader[j].Contains("email"))
                        {
                            reorderData[i, 5] = copyData[i, j];
                        }
                        else
                        {

                        }
                    }
                }
                string thisfile = String.Empty;
                // CREATE A FILE SAVE DIALOG WITH DESIRED FILE FORMAT AND EXTENSION
                SaveFileDialog fileStream = new SaveFileDialog();
                fileStream.FileName = "insightupload.csv";
                fileStream.DefaultExt = ".csv";
                fileStream.Filter = "Comma Separated files (*.csv)|*.csv";
                // DISPLAY THE CREATE FILE SAVE DIALOG BOX TO THE USER
                DialogResult result = fileStream.ShowDialog();
                // OBTAIN THE SAVE FILE NAME/LOCATION FROM USER INPUT  
                if (result == DialogResult.OK)
                {
                    thisfile = fileStream.FileName;
                }
                // CALL CREATE CLASS AND ASSIGN VALUES FOR READ FILE, SAVE FILE, AND PROPERLY-ORDERED DATA
                Create makeFile = new Create(fileType, contactpath, thisfile, reorderData);
                // CALL CREATE CLASS'S CSV-MAKING METHOD
                makeFile.InsightUpload();
                OpenFolder(thisfile);
            }
        }

        private void btnKnowbe4CSV_Click(object sender, EventArgs e)
        {
            string userGroup = txtUserGroup.Text.ToString();
            Read reading = new Read();
            
            // WORD VALUES
            
            if (fileType == "Word")
            {
                reading.NameFile = contactpath;
                // CREATE AN ARRAY TO HOLD HEADER VALUES FROM FILE
                string[] copyHeader = reading.WordTableHeader();
                // CREATE AN ARRAY TO HOLD DATA VALUES FROM FILE
                string[,] copyData = reading.WordDoc();
                // CREATE INDEXES FOR # OF ROWS AND # OF COLUMNS
                int rowcount = copyData.GetUpperBound(0) + 1;
                int colcount = copyData.GetUpperBound(1) + 1;
                // CREATE AN ARRAY TO PLACE DATA VALUES IN DESIRED ORDER
                string[,] reorderData = new string[rowcount, 15];

                try
                {
                    for (int i = 0; i < rowcount; i++)
                    {
                        for (int j = 0; j < colcount; j++)
                        {
                            // BLANK VALUES

                            // SET LOCATION VALUE
                            reorderData[i, 6] = " ";
                            // SET DIVISION VALUE
                            reorderData[i, 7] = " ";
                            // SET MANAGER NAME VALUE
                            reorderData[i, 8] = " ";
                            // SET MANAGER EMAIL VALUE
                            reorderData[i, 9] = " ";
                            // SET EMPLOYEE NUMBER VALUE
                            reorderData[i, 10] = " ";
                            // SET PASSWORD VALUE
                            reorderData[i, 12] = " ";
                            // SET MOBILE NUMBER VALUE
                            reorderData[i, 13] = " ";
                            // SET AD MANAGED VALUE
                            reorderData[i, 14] = " ";
                            // SET USERGROUP VALUE
                            if (!String.IsNullOrWhiteSpace(userGroup))
                            {
                                reorderData[i, 5] = userGroup;
                            }
                            else
                            {
                                reorderData[i, 5] = " ";
                            }
                            // SET EMAIL VALUE
                            if (copyHeader[j].Contains("email"))
                            {
                                reorderData[i, 0] = copyData[i, j];
                            }
                            // SET FIRST NAME VALUE
                            else if (copyHeader[j].Contains("first"))
                            {
                                reorderData[i, 1] = copyData[i, j];
                            }
                            // SET LAST NAME VALUE
                            else if (copyHeader[j].Contains("last"))
                            {
                                reorderData[i, 2] = copyData[i, j];
                            }
                            // SET PHONE NUMBER VALUE
                            else if (copyHeader[j].Contains("phone"))
                            {
                                reorderData[i, 3] = copyData[i, j];
                            }
                            // SET EXTENSION VALUE
                            else if (copyHeader[j].Contains("ext") || copyHeader[j].Contains("extension"))
                            {
                                reorderData[i, 4] = copyData[i, j];
                            }
                            // SET TITLE VALUE
                            else if (copyHeader[j].Contains("title"))
                            {
                                reorderData[i, 11] = copyData[i, j];
                            }
                            else
                            {

                            }
                        }
                    }

                    string thisfile = String.Empty;
                    // CREATE A FILE SAVE DIALOG WITH DESIRED FILE FORMAT AND EXTENSION
                    SaveFileDialog fileStream = new SaveFileDialog();
                    fileStream.FileName = "resellerupload.csv";
                    fileStream.DefaultExt = ".csv";
                    fileStream.Filter = "Comma Separated files (*.csv)|*.csv";
                    // DISPLAY THE CREATE FILE SAVE DIALOG BOX TO THE USER
                    DialogResult result = fileStream.ShowDialog();
                    // OBTAIN THE SAVE FILE NAME/LOCATION FROM USER INPUT  
                    if (result == DialogResult.OK)
                    {
                        thisfile = fileStream.FileName;
                    }
                    // CALL CREATE CLASS AND ASSIGN VALUES FOR READ FILE, SAVE FILE, AND PROPERLY-ORDERED DATA
                    Create makeFile = new Create(fileType, contactpath, thisfile, reorderData);
                    // CALL CREATE CLASS'S CSV-MAKING METHOD
                    makeFile.ResellerUpload();
                    OpenFolder(thisfile);
                }
                catch
                {
                    MessageBox.Show("Unable to export data to file properly.");
                }
            }
            // EXCEL VALUES
            
            else if (fileType == "Excel")
            {
                reading.NameFile = contactpath;
                // CREATE AN ARRAY TO HOLD HEADER VALUES FROM FILE
                string[] copyHeader = reading.ExcelTableHeader();
                // CREATE AN ARRAY TO HOLD DATA VALUES FROM FILE
                string[,] copyData = reading.ExcelDoc();
                // CREATE INDEXES FOR # OF ROWS AND # OF COLUMNS
                int rowcount = copyData.GetUpperBound(0) + 1;
                int colcount = copyData.GetUpperBound(1) + 1;
                // CREATE AN ARRAY TO PLACE DATA VALUES IN DESIRED ORDER
                string[,] reorderData = new string[rowcount, 15];

                try
                {
                    for (int i = 0; i < rowcount; i++)
                    {
                        for (int j = 0; j < colcount; j++)
                        {
                            // BLANK VALUES

                            // SET LOCATION VALUE
                            reorderData[i, 6] = " ";
                            // SET DIVISION VALUE
                            reorderData[i, 7] = " ";
                            // SET MANAGER NAME VALUE
                            reorderData[i, 8] = " ";
                            // SET MANAGER EMAIL VALUE
                            reorderData[i, 9] = " ";
                            // SET EMPLOYEE NUMBER VALUE
                            reorderData[i, 10] = " ";
                            // SET PASSWORD VALUE
                            reorderData[i, 12] = " ";
                            // SET MOBILE NUMBER VALUE
                            reorderData[i, 13] = " ";
                            // SET AD MANAGED VALUE
                            reorderData[i, 14] = " ";
                            // SET USERGROUP VALUE
                            if (!String.IsNullOrWhiteSpace(userGroup))
                            {
                                reorderData[i, 5] = userGroup;
                            }
                            else
                            {
                                reorderData[i, 5] = " ";
                            }
                            // SET EMAIL VALUE
                            if (copyHeader[j].Contains("email"))
                            {
                                reorderData[i, 0] = copyData[i, j];
                            }
                            // SET FIRST NAME VALUE
                            else if (copyHeader[j].Contains("first"))
                            {
                                reorderData[i, 1] = copyData[i, j];
                            }
                            // SET LAST NAME VALUE
                            else if (copyHeader[j].Contains("last"))
                            {
                                reorderData[i, 2] = copyData[i, j];
                            }
                            // SET PHONE NUMBER VALUE
                            else if (copyHeader[j].Contains("phone"))
                            {
                                reorderData[i, 3] = copyData[i, j];
                            }
                            // SET EXTENSION VALUE
                            else if (copyHeader[j].Contains("ext") || copyHeader[j].Contains("extension"))
                            {
                                reorderData[i, 4] = copyData[i, j];
                            }
                            // SET TITLE VALUE
                            else if (copyHeader[j].Contains("title"))
                            {
                                reorderData[i, 11] = copyData[i, j];
                            }
                            else
                            {

                            }
                        }
                    }

                    string thisfile = String.Empty;
                    // CREATE A FILE SAVE DIALOG WITH DESIRED FILE FORMAT AND EXTENSION
                    SaveFileDialog fileStream = new SaveFileDialog();
                    fileStream.FileName = "resellerupload.csv";
                    fileStream.DefaultExt = ".csv";
                    fileStream.Filter = "Comma Separated files (*.csv)|*.csv";
                    // DISPLAY THE CREATE FILE SAVE DIALOG BOX TO THE USER
                    DialogResult result = fileStream.ShowDialog();
                    // OBTAIN THE SAVE FILE NAME/LOCATION FROM USER INPUT  
                    if (result == DialogResult.OK)
                    {
                        thisfile = fileStream.FileName;
                    }
                    // CALL CREATE CLASS AND ASSIGN VALUES FOR READ FILE, SAVE FILE, AND PROPERLY-ORDERED DATA
                    Create makeFile = new Create(fileType, contactpath, thisfile, reorderData);
                    // CALL CREATE CLASS'S CSV-MAKING METHOD
                    makeFile.ResellerUpload();
                    OpenFolder(thisfile);
                }
                catch
                {
                    MessageBox.Show("Unable to export data to file properly.");
                }
            }
        }

        // ####################################################################################################################################################################################################
        // ####################################################################################################################################################################################################
        // ###########################                                                                                                                                              ###########################
        // ###########################                                                                  PAYLOADS                                                                    ###########################
        // ###########################                                                                                                                                              ###########################
        // ####################################################################################################################################################################################################
        // ####################################################################################################################################################################################################


        private void btnCopyClipboard_Click(object sender, EventArgs e)
        {
            if (cboPayloadPicker.SelectedIndex != -1 && radInsight.Checked)
            {
                string[] splitter = cboPayloadPicker.SelectedValue.ToString().Split('\\');
                string filepath = splitter[splitter.Length - 1];
                FileStream payLoad = new FileStream(@"..\\..\\payloads\\" + filepath, FileMode.Open, FileAccess.Read);
                streamer = new StreamReader(payLoad);
                itemText = streamer.ReadToEnd().ToString();
                Clipboard.SetText(itemText);
            }
        }

        private void radInsight_CheckedChanged(object sender, EventArgs e)
        {
            btnCopyClipboard.Enabled = true;
            cboPayloadPicker.Enabled = true;
            btnCopyClipboard.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
            btnCopyClipboard.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(60, 184, 218);
            btnCopyClipboard.ForeColor = System.Drawing.Color.White;
            cboPayloadPicker.BackColor = System.Drawing.Color.FromArgb(50, 60, 70);
            cboPayloadPicker.ForeColor = System.Drawing.Color.White;
            if (radInsight.Checked)
            {
                string key = string.Empty;
                string value = string.Empty;
                int length = path.Length;

                List<KeyValuePair<string, string>> data = new List<KeyValuePair<string, string>>();
                path = @"..\\..\\payloads\\";
                string[] files = Directory.GetFiles(path);
                foreach (var file in files)
                {
                    key = file.ToString();
                    string[] temp = file.ToString().Split('\\');
                    temp = temp[temp.Length-1].Split('.');
                    value = temp[0];
                    data.Add(new KeyValuePair<string, string>(key, value));
                }
                cboPayloadPicker.DataSource = null;
                cboPayloadPicker.Items.Clear();
                cboPayloadPicker.DataSource = new BindingSource(data, null);
                cboPayloadPicker.DisplayMember = "Value";
                cboPayloadPicker.ValueMember = "Key";
            }
        }

        private void btnAddNew_Click(object sender, EventArgs e)
        {
            AddPayload addPayload = new AddPayload();
            addPayload.ShowDialog();
        }


        // ####################################################################################################################################################################################################
        // ####################################################################################################################################################################################################
        // ###########################                                                                                                                                              ###########################
        // ###########################                                                            PHONE CALL TAB                                                                    ###########################
        // ###########################                                                                                                                                              ###########################
        // ####################################################################################################################################################################################################
        // ####################################################################################################################################################################################################

        //method that will create a new Excel Sheet that will be used when making calls to clients 
        private void btnCreateCallList_Click(object sender, EventArgs e)
        {
            NewCallList callList = new NewCallList();
            callList.dataTable = dataTable;
            callList.ShowDialog();
        }

        private void btnMakeCalls_Click(object sender, EventArgs e)
        {
            MakeCalls calls = new MakeCalls();
            if (calls.failed == false)
            {
                calls.ShowDialog();
                calls.driver.Quit();
            }
        }

        // ####################################################################################################################################################################################################
        // ####################################################################################################################################################################################################
        // ###########################                                                                                                                                              ###########################
        // ###########################                                                            CREATE REPORT TAB                                                                 ###########################
        // ###########################                                                                                                                                              ###########################
        // ####################################################################################################################################################################################################
        // ####################################################################################################################################################################################################
        
    }
}
