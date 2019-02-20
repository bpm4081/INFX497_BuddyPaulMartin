namespace INFX497_BuddyPaulMartin
{
    partial class AddPayload
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
            this.rtxtAddPayload = new System.Windows.Forms.RichTextBox();
            this.btnAddPayload = new System.Windows.Forms.Button();
            this.txtPayloadName = new System.Windows.Forms.TextBox();
            this.lblResult = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnImportPayload = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // rtxtAddPayload
            // 
            this.rtxtAddPayload.BackColor = System.Drawing.Color.Gainsboro;
            this.rtxtAddPayload.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.rtxtAddPayload.DetectUrls = false;
            this.rtxtAddPayload.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(60)))), ((int)(((byte)(70)))));
            this.rtxtAddPayload.Location = new System.Drawing.Point(12, 12);
            this.rtxtAddPayload.Name = "rtxtAddPayload";
            this.rtxtAddPayload.Size = new System.Drawing.Size(837, 541);
            this.rtxtAddPayload.TabIndex = 0;
            this.rtxtAddPayload.Text = "";
            this.rtxtAddPayload.TextChanged += new System.EventHandler(this.rtxtAddPayload_TextChanged);
            // 
            // btnAddPayload
            // 
            this.btnAddPayload.BackColor = System.Drawing.Color.Gray;
            this.btnAddPayload.Enabled = false;
            this.btnAddPayload.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(184)))), ((int)(((byte)(218)))));
            this.btnAddPayload.FlatAppearance.BorderSize = 2;
            this.btnAddPayload.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAddPayload.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAddPayload.ForeColor = System.Drawing.Color.LightGray;
            this.btnAddPayload.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnAddPayload.ImageIndex = 7;
            this.btnAddPayload.Location = new System.Drawing.Point(497, 559);
            this.btnAddPayload.Name = "btnAddPayload";
            this.btnAddPayload.Size = new System.Drawing.Size(129, 65);
            this.btnAddPayload.TabIndex = 4;
            this.btnAddPayload.Text = "Add New Payload";
            this.btnAddPayload.UseVisualStyleBackColor = false;
            this.btnAddPayload.Click += new System.EventHandler(this.btnAddPayload_Click);
            // 
            // txtPayloadName
            // 
            this.txtPayloadName.Location = new System.Drawing.Point(160, 604);
            this.txtPayloadName.Name = "txtPayloadName";
            this.txtPayloadName.Size = new System.Drawing.Size(305, 20);
            this.txtPayloadName.TabIndex = 7;
            this.txtPayloadName.TextChanged += new System.EventHandler(this.txtPayloadName_TextChanged);
            // 
            // lblResult
            // 
            this.lblResult.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblResult.Location = new System.Drawing.Point(646, 559);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(203, 65);
            this.lblResult.TabIndex = 15;
            this.lblResult.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.Control;
            this.label1.Location = new System.Drawing.Point(156, 571);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(309, 20);
            this.label1.TabIndex = 16;
            this.label1.Text = "Payload Name without File Extension:";
            // 
            // btnImportPayload
            // 
            this.btnImportPayload.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(60)))), ((int)(((byte)(70)))));
            this.btnImportPayload.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(184)))), ((int)(((byte)(218)))));
            this.btnImportPayload.FlatAppearance.BorderSize = 2;
            this.btnImportPayload.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnImportPayload.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnImportPayload.ForeColor = System.Drawing.Color.White;
            this.btnImportPayload.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnImportPayload.ImageKey = "add-file-256.gif";
            this.btnImportPayload.Location = new System.Drawing.Point(12, 559);
            this.btnImportPayload.Name = "btnImportPayload";
            this.btnImportPayload.Size = new System.Drawing.Size(126, 65);
            this.btnImportPayload.TabIndex = 17;
            this.btnImportPayload.Text = "Import Payload";
            this.btnImportPayload.UseVisualStyleBackColor = false;
            this.btnImportPayload.Click += new System.EventHandler(this.btnImportPayload_Click);
            // 
            // AddPayload
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(60)))), ((int)(((byte)(70)))));
            this.ClientSize = new System.Drawing.Size(861, 636);
            this.Controls.Add(this.btnImportPayload);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.txtPayloadName);
            this.Controls.Add(this.btnAddPayload);
            this.Controls.Add(this.rtxtAddPayload);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "AddPayload";
            this.Text = "Add a Payload";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox rtxtAddPayload;
        private System.Windows.Forms.Button btnAddPayload;
        private System.Windows.Forms.TextBox txtPayloadName;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnImportPayload;
    }
}