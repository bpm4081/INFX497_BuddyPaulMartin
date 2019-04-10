﻿namespace INFX497_BuddyPaulMartin
{
    partial class Main
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.tabMAIN = new System.Windows.Forms.TabControl();
            this.tabEmail = new System.Windows.Forms.TabPage();
            this.btnAddNew = new System.Windows.Forms.Button();
            this.imagesSmall = new System.Windows.Forms.ImageList(this.components);
            this.label4 = new System.Windows.Forms.Label();
            this.txtUserGroup = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.radInsight = new System.Windows.Forms.RadioButton();
            this.btnCopyClipboard = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.cboPayloadPicker = new System.Windows.Forms.ComboBox();
            this.btnInsightCSV = new System.Windows.Forms.Button();
            this.btnKnowbe4CSV = new System.Windows.Forms.Button();
            this.tabPhone = new System.Windows.Forms.TabPage();
            this.btnMakeCalls = new System.Windows.Forms.Button();
            this.btnCreateCallList = new System.Windows.Forms.Button();
            this.tabReport = new System.Windows.Forms.TabPage();
            this.labelCurrentAction = new System.Windows.Forms.Label();
            this.labelPercentage = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtPOC = new System.Windows.Forms.TextBox();
            this.txtClient = new System.Windows.Forms.TextBox();
            this.radBoth = new System.Windows.Forms.RadioButton();
            this.radPhone = new System.Windows.Forms.RadioButton();
            this.radEmail = new System.Windows.Forms.RadioButton();
            this.label2 = new System.Windows.Forms.Label();
            this.btnReportShell = new System.Windows.Forms.Button();
            this.imagesMedium = new System.Windows.Forms.ImageList(this.components);
            this.panelHead = new System.Windows.Forms.Panel();
            this.btnPreview2 = new System.Windows.Forms.Button();
            this.btnOpenExcelFile = new System.Windows.Forms.Button();
            this.imagesLarge = new System.Windows.Forms.ImageList(this.components);
            this.btnPreview = new System.Windows.Forms.Button();
            this.lblPullContacts = new System.Windows.Forms.Label();
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.createVishingReport = new System.ComponentModel.BackgroundWorker();
            this.createPhishingReport = new System.ComponentModel.BackgroundWorker();
            this.createBothReport = new System.ComponentModel.BackgroundWorker();
            this.tabMAIN.SuspendLayout();
            this.tabEmail.SuspendLayout();
            this.tabPhone.SuspendLayout();
            this.tabReport.SuspendLayout();
            this.panelHead.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMAIN
            // 
            this.tabMAIN.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.tabMAIN.Controls.Add(this.tabEmail);
            this.tabMAIN.Controls.Add(this.tabPhone);
            this.tabMAIN.Controls.Add(this.tabReport);
            this.tabMAIN.HotTrack = true;
            this.tabMAIN.ImageList = this.imagesMedium;
            this.tabMAIN.ItemSize = new System.Drawing.Size(70, 40);
            this.tabMAIN.Location = new System.Drawing.Point(0, 134);
            this.tabMAIN.Margin = new System.Windows.Forms.Padding(0);
            this.tabMAIN.Multiline = true;
            this.tabMAIN.Name = "tabMAIN";
            this.tabMAIN.RightToLeftLayout = true;
            this.tabMAIN.SelectedIndex = 0;
            this.tabMAIN.ShowToolTips = true;
            this.tabMAIN.Size = new System.Drawing.Size(540, 289);
            this.tabMAIN.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.tabMAIN.TabIndex = 9;
            // 
            // tabEmail
            // 
            this.tabEmail.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(60)))), ((int)(((byte)(70)))));
            this.tabEmail.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tabEmail.Controls.Add(this.btnAddNew);
            this.tabEmail.Controls.Add(this.label4);
            this.tabEmail.Controls.Add(this.txtUserGroup);
            this.tabEmail.Controls.Add(this.label1);
            this.tabEmail.Controls.Add(this.radInsight);
            this.tabEmail.Controls.Add(this.btnCopyClipboard);
            this.tabEmail.Controls.Add(this.label3);
            this.tabEmail.Controls.Add(this.cboPayloadPicker);
            this.tabEmail.Controls.Add(this.btnInsightCSV);
            this.tabEmail.Controls.Add(this.btnKnowbe4CSV);
            this.tabEmail.ImageIndex = 2;
            this.tabEmail.Location = new System.Drawing.Point(4, 44);
            this.tabEmail.Margin = new System.Windows.Forms.Padding(4);
            this.tabEmail.Name = "tabEmail";
            this.tabEmail.Padding = new System.Windows.Forms.Padding(4);
            this.tabEmail.Size = new System.Drawing.Size(532, 241);
            this.tabEmail.TabIndex = 0;
            this.tabEmail.ToolTipText = "Email Tools";
            // 
            // btnAddNew
            // 
            this.btnAddNew.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(60)))), ((int)(((byte)(70)))));
            this.btnAddNew.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(184)))), ((int)(((byte)(218)))));
            this.btnAddNew.FlatAppearance.BorderSize = 2;
            this.btnAddNew.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAddNew.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAddNew.ForeColor = System.Drawing.Color.White;
            this.btnAddNew.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAddNew.ImageKey = "add-file-256.gif";
            this.btnAddNew.ImageList = this.imagesSmall;
            this.btnAddNew.Location = new System.Drawing.Point(356, 165);
            this.btnAddNew.Margin = new System.Windows.Forms.Padding(4);
            this.btnAddNew.Name = "btnAddNew";
            this.btnAddNew.Size = new System.Drawing.Size(133, 44);
            this.btnAddNew.TabIndex = 14;
            this.btnAddNew.Text = "Add New";
            this.btnAddNew.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnAddNew.UseVisualStyleBackColor = false;
            this.btnAddNew.Click += new System.EventHandler(this.btnAddNew_Click);
            // 
            // imagesSmall
            // 
            this.imagesSmall.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imagesSmall.ImageStream")));
            this.imagesSmall.TransparentColor = System.Drawing.Color.Transparent;
            this.imagesSmall.Images.SetKeyName(0, "arrow-59-256.gif");
            this.imagesSmall.Images.SetKeyName(1, "add-file-256.gif");
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label4.Location = new System.Drawing.Point(20, 165);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(160, 16);
            this.label4.TabIndex = 13;
            this.label4.Text = "optional user group name";
            // 
            // txtUserGroup
            // 
            this.txtUserGroup.BackColor = System.Drawing.Color.Gray;
            this.txtUserGroup.Enabled = false;
            this.txtUserGroup.Location = new System.Drawing.Point(37, 185);
            this.txtUserGroup.Margin = new System.Windows.Forms.Padding(4);
            this.txtUserGroup.Name = "txtUserGroup";
            this.txtUserGroup.Size = new System.Drawing.Size(123, 22);
            this.txtUserGroup.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label1.Location = new System.Drawing.Point(71, 15);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 25);
            this.label1.TabIndex = 6;
            this.label1.Text = "CSV";
            // 
            // radInsight
            // 
            this.radInsight.AutoSize = true;
            this.radInsight.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radInsight.ForeColor = System.Drawing.Color.White;
            this.radInsight.Location = new System.Drawing.Point(332, 52);
            this.radInsight.Margin = new System.Windows.Forms.Padding(4);
            this.radInsight.Name = "radInsight";
            this.radInsight.Size = new System.Drawing.Size(151, 21);
            this.radInsight.TabIndex = 11;
            this.radInsight.TabStop = true;
            this.radInsight.Text = "Access Payloads";
            this.radInsight.UseVisualStyleBackColor = true;
            this.radInsight.CheckedChanged += new System.EventHandler(this.radInsight_CheckedChanged);
            // 
            // btnCopyClipboard
            // 
            this.btnCopyClipboard.BackColor = System.Drawing.Color.Gray;
            this.btnCopyClipboard.Enabled = false;
            this.btnCopyClipboard.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnCopyClipboard.FlatAppearance.BorderSize = 2;
            this.btnCopyClipboard.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCopyClipboard.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCopyClipboard.ForeColor = System.Drawing.Color.LightGray;
            this.btnCopyClipboard.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCopyClipboard.ImageKey = "arrow-59-256.gif";
            this.btnCopyClipboard.ImageList = this.imagesSmall;
            this.btnCopyClipboard.Location = new System.Drawing.Point(356, 113);
            this.btnCopyClipboard.Margin = new System.Windows.Forms.Padding(4);
            this.btnCopyClipboard.Name = "btnCopyClipboard";
            this.btnCopyClipboard.Size = new System.Drawing.Size(133, 44);
            this.btnCopyClipboard.TabIndex = 9;
            this.btnCopyClipboard.Text = "Clipboard";
            this.btnCopyClipboard.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnCopyClipboard.UseVisualStyleBackColor = false;
            this.btnCopyClipboard.Click += new System.EventHandler(this.btnCopyClipboard_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label3.Location = new System.Drawing.Point(291, 15);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(156, 25);
            this.label3.TabIndex = 8;
            this.label3.Text = "Payload Picker";
            // 
            // cboPayloadPicker
            // 
            this.cboPayloadPicker.BackColor = System.Drawing.Color.Gray;
            this.cboPayloadPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboPayloadPicker.Enabled = false;
            this.cboPayloadPicker.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cboPayloadPicker.ForeColor = System.Drawing.Color.LightGray;
            this.cboPayloadPicker.FormattingEnabled = true;
            this.cboPayloadPicker.Location = new System.Drawing.Point(257, 80);
            this.cboPayloadPicker.Margin = new System.Windows.Forms.Padding(4);
            this.cboPayloadPicker.MaxDropDownItems = 20;
            this.cboPayloadPicker.Name = "cboPayloadPicker";
            this.cboPayloadPicker.Size = new System.Drawing.Size(235, 24);
            this.cboPayloadPicker.TabIndex = 6;
            // 
            // btnInsightCSV
            // 
            this.btnInsightCSV.BackColor = System.Drawing.Color.Gray;
            this.btnInsightCSV.Enabled = false;
            this.btnInsightCSV.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnInsightCSV.FlatAppearance.BorderSize = 2;
            this.btnInsightCSV.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnInsightCSV.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnInsightCSV.ForeColor = System.Drawing.Color.LightGray;
            this.btnInsightCSV.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnInsightCSV.ImageIndex = 7;
            this.btnInsightCSV.Location = new System.Drawing.Point(37, 53);
            this.btnInsightCSV.Margin = new System.Windows.Forms.Padding(4);
            this.btnInsightCSV.Name = "btnInsightCSV";
            this.btnInsightCSV.Size = new System.Drawing.Size(124, 43);
            this.btnInsightCSV.TabIndex = 3;
            this.btnInsightCSV.Text = "Insight";
            this.btnInsightCSV.UseVisualStyleBackColor = false;
            this.btnInsightCSV.Click += new System.EventHandler(this.btnInsightCSV_Click);
            // 
            // btnKnowbe4CSV
            // 
            this.btnKnowbe4CSV.BackColor = System.Drawing.Color.Gray;
            this.btnKnowbe4CSV.Enabled = false;
            this.btnKnowbe4CSV.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btnKnowbe4CSV.FlatAppearance.BorderSize = 2;
            this.btnKnowbe4CSV.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnKnowbe4CSV.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnKnowbe4CSV.ForeColor = System.Drawing.SystemColors.ScrollBar;
            this.btnKnowbe4CSV.Location = new System.Drawing.Point(37, 113);
            this.btnKnowbe4CSV.Margin = new System.Windows.Forms.Padding(4);
            this.btnKnowbe4CSV.Name = "btnKnowbe4CSV";
            this.btnKnowbe4CSV.Size = new System.Drawing.Size(124, 43);
            this.btnKnowbe4CSV.TabIndex = 4;
            this.btnKnowbe4CSV.Text = "KnowBe4";
            this.btnKnowbe4CSV.UseVisualStyleBackColor = false;
            this.btnKnowbe4CSV.Click += new System.EventHandler(this.btnKnowbe4CSV_Click);
            // 
            // tabPhone
            // 
            this.tabPhone.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(60)))), ((int)(((byte)(70)))));
            this.tabPhone.Controls.Add(this.btnMakeCalls);
            this.tabPhone.Controls.Add(this.btnCreateCallList);
            this.tabPhone.ImageIndex = 8;
            this.tabPhone.Location = new System.Drawing.Point(4, 44);
            this.tabPhone.Margin = new System.Windows.Forms.Padding(4);
            this.tabPhone.Name = "tabPhone";
            this.tabPhone.Padding = new System.Windows.Forms.Padding(4);
            this.tabPhone.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.tabPhone.Size = new System.Drawing.Size(532, 241);
            this.tabPhone.TabIndex = 1;
            this.tabPhone.ToolTipText = "Phone Tools";
            // 
            // btnMakeCalls
            // 
            this.btnMakeCalls.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(60)))), ((int)(((byte)(70)))));
            this.btnMakeCalls.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(184)))), ((int)(((byte)(218)))));
            this.btnMakeCalls.FlatAppearance.BorderSize = 2;
            this.btnMakeCalls.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMakeCalls.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMakeCalls.ForeColor = System.Drawing.Color.White;
            this.btnMakeCalls.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnMakeCalls.ImageIndex = 7;
            this.btnMakeCalls.Location = new System.Drawing.Point(312, 52);
            this.btnMakeCalls.Margin = new System.Windows.Forms.Padding(4);
            this.btnMakeCalls.Name = "btnMakeCalls";
            this.btnMakeCalls.Size = new System.Drawing.Size(124, 54);
            this.btnMakeCalls.TabIndex = 5;
            this.btnMakeCalls.Text = "Make Calls";
            this.btnMakeCalls.UseVisualStyleBackColor = false;
            this.btnMakeCalls.Click += new System.EventHandler(this.btnMakeCalls_Click);
            // 
            // btnCreateCallList
            // 
            this.btnCreateCallList.BackColor = System.Drawing.Color.Gray;
            this.btnCreateCallList.Enabled = false;
            this.btnCreateCallList.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(184)))), ((int)(((byte)(218)))));
            this.btnCreateCallList.FlatAppearance.BorderSize = 2;
            this.btnCreateCallList.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCreateCallList.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateCallList.ForeColor = System.Drawing.Color.LightGray;
            this.btnCreateCallList.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnCreateCallList.ImageIndex = 7;
            this.btnCreateCallList.Location = new System.Drawing.Point(83, 52);
            this.btnCreateCallList.Margin = new System.Windows.Forms.Padding(4);
            this.btnCreateCallList.Name = "btnCreateCallList";
            this.btnCreateCallList.Size = new System.Drawing.Size(124, 54);
            this.btnCreateCallList.TabIndex = 4;
            this.btnCreateCallList.Text = "Create Call List";
            this.btnCreateCallList.UseVisualStyleBackColor = false;
            this.btnCreateCallList.Click += new System.EventHandler(this.btnCreateCallList_Click);
            // 
            // tabReport
            // 
            this.tabReport.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(60)))), ((int)(((byte)(70)))));
            this.tabReport.Controls.Add(this.progressBar1);
            this.tabReport.Controls.Add(this.labelCurrentAction);
            this.tabReport.Controls.Add(this.labelPercentage);
            this.tabReport.Controls.Add(this.dateTimePicker1);
            this.tabReport.Controls.Add(this.label7);
            this.tabReport.Controls.Add(this.label6);
            this.tabReport.Controls.Add(this.label5);
            this.tabReport.Controls.Add(this.txtPOC);
            this.tabReport.Controls.Add(this.txtClient);
            this.tabReport.Controls.Add(this.radBoth);
            this.tabReport.Controls.Add(this.radPhone);
            this.tabReport.Controls.Add(this.radEmail);
            this.tabReport.Controls.Add(this.label2);
            this.tabReport.Controls.Add(this.btnReportShell);
            this.tabReport.ImageIndex = 0;
            this.tabReport.Location = new System.Drawing.Point(4, 44);
            this.tabReport.Margin = new System.Windows.Forms.Padding(4);
            this.tabReport.Name = "tabReport";
            this.tabReport.Padding = new System.Windows.Forms.Padding(4);
            this.tabReport.Size = new System.Drawing.Size(532, 241);
            this.tabReport.TabIndex = 2;
            // 
            // labelCurrentAction
            // 
            this.labelCurrentAction.AutoSize = true;
            this.labelCurrentAction.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelCurrentAction.ForeColor = System.Drawing.Color.White;
            this.labelCurrentAction.Location = new System.Drawing.Point(213, 90);
            this.labelCurrentAction.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelCurrentAction.Name = "labelCurrentAction";
            this.labelCurrentAction.Size = new System.Drawing.Size(103, 24);
            this.labelCurrentAction.TabIndex = 22;
            this.labelCurrentAction.Text = "Loading...";
            this.labelCurrentAction.Visible = false;
            // 
            // labelPercentage
            // 
            this.labelPercentage.AutoSize = true;
            this.labelPercentage.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelPercentage.ForeColor = System.Drawing.Color.White;
            this.labelPercentage.Location = new System.Drawing.Point(133, 150);
            this.labelPercentage.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelPercentage.Name = "labelPercentage";
            this.labelPercentage.Size = new System.Drawing.Size(35, 20);
            this.labelPercentage.TabIndex = 21;
            this.labelPercentage.Text = "0%";
            this.labelPercentage.Visible = false;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(131, 118);
            this.progressBar1.Margin = new System.Windows.Forms.Padding(4);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(267, 28);
            this.progressBar1.TabIndex = 20;
            this.progressBar1.Visible = false;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker1.Location = new System.Drawing.Point(189, 122);
            this.dateTimePicker1.Margin = new System.Windows.Forms.Padding(4);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(132, 22);
            this.dateTimePicker1.TabIndex = 17;
            this.dateTimePicker1.Value = new System.DateTime(2019, 1, 1, 0, 0, 0, 0);
            this.dateTimePicker1.ValueChanged += new System.EventHandler(this.dateTimePicker1_ValueChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label7.Location = new System.Drawing.Point(5, 90);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(164, 24);
            this.label7.TabIndex = 19;
            this.label7.Text = "Point of Contact*";
            this.label7.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label6.Location = new System.Drawing.Point(5, 124);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(157, 24);
            this.label6.TabIndex = 18;
            this.label6.Text = "Email Start Date";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label5.Location = new System.Drawing.Point(75, 55);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(106, 24);
            this.label5.TabIndex = 17;
            this.label5.Text = "Company*";
            // 
            // txtPOC
            // 
            this.txtPOC.Location = new System.Drawing.Point(189, 89);
            this.txtPOC.Margin = new System.Windows.Forms.Padding(5);
            this.txtPOC.Name = "txtPOC";
            this.txtPOC.Size = new System.Drawing.Size(132, 22);
            this.txtPOC.TabIndex = 16;
            this.txtPOC.TextChanged += new System.EventHandler(this.txtPOC_TextChanged);
            // 
            // txtClient
            // 
            this.txtClient.Location = new System.Drawing.Point(189, 54);
            this.txtClient.Margin = new System.Windows.Forms.Padding(5);
            this.txtClient.Name = "txtClient";
            this.txtClient.Size = new System.Drawing.Size(132, 22);
            this.txtClient.TabIndex = 15;
            this.txtClient.TextChanged += new System.EventHandler(this.txtClient_TextChanged);
            // 
            // radBoth
            // 
            this.radBoth.AutoSize = true;
            this.radBoth.Enabled = false;
            this.radBoth.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radBoth.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.radBoth.Location = new System.Drawing.Point(377, 127);
            this.radBoth.Margin = new System.Windows.Forms.Padding(1);
            this.radBoth.Name = "radBoth";
            this.radBoth.Size = new System.Drawing.Size(62, 21);
            this.radBoth.TabIndex = 13;
            this.radBoth.TabStop = true;
            this.radBoth.Text = "Both";
            this.radBoth.UseVisualStyleBackColor = true;
            this.radBoth.Click += new System.EventHandler(this.radBoth_Click);
            // 
            // radPhone
            // 
            this.radPhone.AutoSize = true;
            this.radPhone.Enabled = false;
            this.radPhone.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radPhone.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.radPhone.Location = new System.Drawing.Point(377, 92);
            this.radPhone.Margin = new System.Windows.Forms.Padding(1);
            this.radPhone.Name = "radPhone";
            this.radPhone.Size = new System.Drawing.Size(75, 21);
            this.radPhone.TabIndex = 12;
            this.radPhone.TabStop = true;
            this.radPhone.Text = "Phone";
            this.radPhone.UseVisualStyleBackColor = true;
            this.radPhone.Click += new System.EventHandler(this.radPhone_Click);
            // 
            // radEmail
            // 
            this.radEmail.AutoSize = true;
            this.radEmail.Enabled = false;
            this.radEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radEmail.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.radEmail.Location = new System.Drawing.Point(377, 58);
            this.radEmail.Margin = new System.Windows.Forms.Padding(1);
            this.radEmail.Name = "radEmail";
            this.radEmail.Size = new System.Drawing.Size(68, 21);
            this.radEmail.TabIndex = 11;
            this.radEmail.TabStop = true;
            this.radEmail.Text = "Email";
            this.radEmail.UseVisualStyleBackColor = true;
            this.radEmail.Click += new System.EventHandler(this.radEmail_Click);
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label2.Location = new System.Drawing.Point(188, 15);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(135, 21);
            this.label2.TabIndex = 9;
            this.label2.Text = "Report Shell";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnReportShell
            // 
            this.btnReportShell.BackColor = System.Drawing.Color.Gray;
            this.btnReportShell.Enabled = false;
            this.btnReportShell.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(184)))), ((int)(((byte)(218)))));
            this.btnReportShell.FlatAppearance.BorderSize = 2;
            this.btnReportShell.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnReportShell.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnReportShell.ForeColor = System.Drawing.Color.LightGray;
            this.btnReportShell.Location = new System.Drawing.Point(192, 162);
            this.btnReportShell.Margin = new System.Windows.Forms.Padding(5);
            this.btnReportShell.Name = "btnReportShell";
            this.btnReportShell.Size = new System.Drawing.Size(135, 63);
            this.btnReportShell.TabIndex = 8;
            this.btnReportShell.Text = "Create Report";
            this.btnReportShell.UseVisualStyleBackColor = false;
            this.btnReportShell.Click += new System.EventHandler(this.btnReportShell_Click);
            // 
            // imagesMedium
            // 
            this.imagesMedium.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imagesMedium.ImageStream")));
            this.imagesMedium.TransparentColor = System.Drawing.Color.Transparent;
            this.imagesMedium.Images.SetKeyName(0, "text-file-6-512.ico");
            this.imagesMedium.Images.SetKeyName(1, "phone-18-256.ico");
            this.imagesMedium.Images.SetKeyName(2, "envelope-open-512.ico");
            this.imagesMedium.Images.SetKeyName(3, "visible-512.ico");
            this.imagesMedium.Images.SetKeyName(4, "microsoft-word-512.ico");
            this.imagesMedium.Images.SetKeyName(5, "csv-256.gif");
            this.imagesMedium.Images.SetKeyName(6, "arrow-53-256.gif");
            this.imagesMedium.Images.SetKeyName(7, "excel-3-512.ico");
            this.imagesMedium.Images.SetKeyName(8, "phone-2-512.ico");
            // 
            // panelHead
            // 
            this.panelHead.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(184)))), ((int)(((byte)(218)))));
            this.panelHead.Controls.Add(this.btnPreview2);
            this.panelHead.Controls.Add(this.btnOpenExcelFile);
            this.panelHead.Controls.Add(this.btnPreview);
            this.panelHead.Controls.Add(this.lblPullContacts);
            this.panelHead.Controls.Add(this.btnOpenFile);
            this.panelHead.Location = new System.Drawing.Point(0, 0);
            this.panelHead.Margin = new System.Windows.Forms.Padding(0);
            this.panelHead.Name = "panelHead";
            this.panelHead.Size = new System.Drawing.Size(540, 134);
            this.panelHead.TabIndex = 8;
            // 
            // btnPreview2
            // 
            this.btnPreview2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(184)))), ((int)(((byte)(218)))));
            this.btnPreview2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnPreview2.BackgroundImage")));
            this.btnPreview2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnPreview2.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnPreview2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPreview2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(60)))), ((int)(((byte)(70)))));
            this.btnPreview2.Location = new System.Drawing.Point(317, 15);
            this.btnPreview2.Margin = new System.Windows.Forms.Padding(1);
            this.btnPreview2.Name = "btnPreview2";
            this.btnPreview2.Size = new System.Drawing.Size(36, 27);
            this.btnPreview2.TabIndex = 3;
            this.btnPreview2.UseVisualStyleBackColor = false;
            this.btnPreview2.Visible = false;
            this.btnPreview2.Click += new System.EventHandler(this.btnPreview2_Click);
            // 
            // btnOpenExcelFile
            // 
            this.btnOpenExcelFile.FlatAppearance.BorderColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnOpenExcelFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOpenExcelFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOpenExcelFile.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnOpenExcelFile.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnOpenExcelFile.ImageIndex = 1;
            this.btnOpenExcelFile.ImageList = this.imagesLarge;
            this.btnOpenExcelFile.Location = new System.Drawing.Point(359, 15);
            this.btnOpenExcelFile.Margin = new System.Windows.Forms.Padding(4, 4, 27, 4);
            this.btnOpenExcelFile.Name = "btnOpenExcelFile";
            this.btnOpenExcelFile.Size = new System.Drawing.Size(139, 68);
            this.btnOpenExcelFile.TabIndex = 2;
            this.btnOpenExcelFile.Text = "Open\r\nExcel\r\nFile";
            this.btnOpenExcelFile.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnOpenExcelFile.UseVisualStyleBackColor = true;
            this.btnOpenExcelFile.Click += new System.EventHandler(this.btnOpenExcelFile_Click);
            // 
            // imagesLarge
            // 
            this.imagesLarge.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imagesLarge.ImageStream")));
            this.imagesLarge.TransparentColor = System.Drawing.Color.Transparent;
            this.imagesLarge.Images.SetKeyName(0, "microsoft-word-512.gif");
            this.imagesLarge.Images.SetKeyName(1, "excel-3-512.gif");
            this.imagesLarge.Images.SetKeyName(2, "envelope-closed-512.ico");
            this.imagesLarge.Images.SetKeyName(3, "document-2-512.ico");
            this.imagesLarge.Images.SetKeyName(4, "phone-70-512.gif");
            // 
            // btnPreview
            // 
            this.btnPreview.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(184)))), ((int)(((byte)(218)))));
            this.btnPreview.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnPreview.BackgroundImage")));
            this.btnPreview.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnPreview.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnPreview.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPreview.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(60)))), ((int)(((byte)(70)))));
            this.btnPreview.Location = new System.Drawing.Point(176, 15);
            this.btnPreview.Margin = new System.Windows.Forms.Padding(1);
            this.btnPreview.Name = "btnPreview";
            this.btnPreview.Size = new System.Drawing.Size(36, 27);
            this.btnPreview.TabIndex = 0;
            this.btnPreview.UseVisualStyleBackColor = false;
            this.btnPreview.Visible = false;
            this.btnPreview.Click += new System.EventHandler(this.btnPreview_Click);
            // 
            // lblPullContacts
            // 
            this.lblPullContacts.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.999999F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPullContacts.Location = new System.Drawing.Point(4, 86);
            this.lblPullContacts.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblPullContacts.Name = "lblPullContacts";
            this.lblPullContacts.Size = new System.Drawing.Size(251, 50);
            this.lblPullContacts.TabIndex = 1;
            this.lblPullContacts.Text = "Contact Data Extracted!";
            this.lblPullContacts.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblPullContacts.Visible = false;
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.FlatAppearance.BorderColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnOpenFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOpenFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOpenFile.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnOpenFile.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnOpenFile.ImageIndex = 0;
            this.btnOpenFile.ImageList = this.imagesLarge;
            this.btnOpenFile.Location = new System.Drawing.Point(27, 15);
            this.btnOpenFile.Margin = new System.Windows.Forms.Padding(27, 4, 4, 4);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(139, 68);
            this.btnOpenFile.TabIndex = 0;
            this.btnOpenFile.Text = "Open \r\nWord\r\nFile";
            this.btnOpenFile.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // toolTips
            // 
            this.toolTips.ToolTipTitle = "Email Tools";
            // 
            // createVishingReport
            // 
            this.createVishingReport.WorkerReportsProgress = true;
            this.createVishingReport.WorkerSupportsCancellation = true;
            this.createVishingReport.DoWork += new System.ComponentModel.DoWorkEventHandler(this.createVishingReport_DoWork);
            this.createVishingReport.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.createVishingReport_ProgressChanged);
            this.createVishingReport.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.createVishingReport_RunWorkerCompleted);
            // 
            // createPhishingReport
            // 
            this.createPhishingReport.WorkerReportsProgress = true;
            this.createPhishingReport.WorkerSupportsCancellation = true;
            this.createPhishingReport.DoWork += new System.ComponentModel.DoWorkEventHandler(this.createPhishingReport_DoWork);
            this.createPhishingReport.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.createPhishingReport_ProgressChanged);
            this.createPhishingReport.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.createPhishingReport_RunWorkerCompleted);
            // 
            // createBothReport
            // 
            this.createBothReport.WorkerReportsProgress = true;
            this.createBothReport.WorkerSupportsCancellation = true;
            this.createBothReport.DoWork += new System.ComponentModel.DoWorkEventHandler(this.createBothReport_DoWork);
            this.createBothReport.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.createBothReport_ProgressChanged);
            this.createBothReport.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.createBothReport_RunWorkerCompleted);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(60)))), ((int)(((byte)(70)))));
            this.ClientSize = new System.Drawing.Size(536, 425);
            this.Controls.Add(this.tabMAIN);
            this.Controls.Add(this.panelHead);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.IsMdiContainer = true;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Main";
            this.Text = "Trace Remote Social Engineering Tool";
            this.Load += new System.EventHandler(this.Main_Load);
            this.tabMAIN.ResumeLayout(false);
            this.tabEmail.ResumeLayout(false);
            this.tabEmail.PerformLayout();
            this.tabPhone.ResumeLayout(false);
            this.tabReport.ResumeLayout(false);
            this.tabReport.PerformLayout();
            this.panelHead.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabMAIN;
        private System.Windows.Forms.TabPage tabEmail;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtUserGroup;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton radInsight;
        private System.Windows.Forms.Button btnCopyClipboard;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cboPayloadPicker;
        private System.Windows.Forms.Button btnInsightCSV;
        private System.Windows.Forms.Button btnKnowbe4CSV;
        private System.Windows.Forms.TabPage tabPhone;
        private System.Windows.Forms.Panel panelHead;
        private System.Windows.Forms.Label lblPullContacts;
        private System.Windows.Forms.Button btnOpenFile;
        public System.Windows.Forms.ImageList imagesMedium;
        private System.Windows.Forms.Button btnPreview;
        private System.Windows.Forms.ToolTip toolTips;
        private System.Windows.Forms.TabPage tabReport;
        private System.Windows.Forms.RadioButton radBoth;
        private System.Windows.Forms.RadioButton radPhone;
        private System.Windows.Forms.RadioButton radEmail;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnReportShell;
        private System.Windows.Forms.ImageList imagesSmall;
        private System.Windows.Forms.Button btnAddNew;
        private System.Windows.Forms.TextBox txtPOC;
        private System.Windows.Forms.TextBox txtClient;
        private System.Windows.Forms.Button btnOpenExcelFile;
        private System.Windows.Forms.Button btnPreview2;
        private System.Windows.Forms.Button btnCreateCallList;
        private System.Windows.Forms.ImageList imagesLarge;
        private System.Windows.Forms.Button btnMakeCalls;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label labelPercentage;
        private System.Windows.Forms.Label labelCurrentAction;
        private System.ComponentModel.BackgroundWorker createVishingReport;
        private System.ComponentModel.BackgroundWorker createPhishingReport;
        private System.ComponentModel.BackgroundWorker createBothReport;
    }
}

