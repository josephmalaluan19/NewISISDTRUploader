namespace NewUploader
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.ExcelUploader = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.btnUploadExcel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ExcelUploader
            // 
            this.ExcelUploader.Location = new System.Drawing.Point(310, 27);
            this.ExcelUploader.Name = "ExcelUploader";
            this.ExcelUploader.Size = new System.Drawing.Size(223, 23);
            this.ExcelUploader.TabIndex = 0;
            this.ExcelUploader.Text = "Browse Excel";
            this.ExcelUploader.Click += new System.EventHandler(this.ExcelUploader_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 15);
            this.label1.TabIndex = 1;
            this.label1.Text = "label1";
            // 
            // btnUploadExcel
            // 
            this.btnUploadExcel.Location = new System.Drawing.Point(539, 27);
            this.btnUploadExcel.Name = "btnUploadExcel";
            this.btnUploadExcel.Size = new System.Drawing.Size(190, 23);
            this.btnUploadExcel.TabIndex = 2;
            this.btnUploadExcel.Text = "Upload Excel";
            this.btnUploadExcel.UseVisualStyleBackColor = true;
            this.btnUploadExcel.Click += new System.EventHandler(this.btnUploadExcel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnUploadExcel);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ExcelUploader);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private Button ExcelUploader;
        private OpenFileDialog openFileDialog1;
        private Label label1;
        private Button btnUploadExcel;
    }
}