namespace SchedulingAssistant
{
    partial class SchedulingAssistantForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SchedulingAssistantForm));
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.SelectFileButton = new System.Windows.Forms.Button();
            this.ImportFileButton = new System.Windows.Forms.Button();
            this.ExportFileButton = new System.Windows.Forms.Button();
            this.AddRowButton = new System.Windows.Forms.Button();
            this.DeleteRowButton = new System.Windows.Forms.Button();
            this.CheckConflictButton = new System.Windows.Forms.Button();
            this.excelFilePathTB = new System.Windows.Forms.TextBox();
            this.ExcelSheetNames = new System.Windows.Forms.ComboBox();
            this.SheetNamesLabel = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.GenerateReportButton = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.deleteEmptyColumns = new System.Windows.Forms.Button();
            this.deleteEmptyRows = new System.Windows.Forms.Button();
            this.Images = new System.Windows.Forms.ImageList(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView
            // 
            this.dataGridView.BackgroundColor = System.Drawing.Color.WhiteSmoke;
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Location = new System.Drawing.Point(12, 146);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.Size = new System.Drawing.Size(1344, 497);
            this.dataGridView.TabIndex = 0;
            this.dataGridView.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_CellEnter);
            this.dataGridView.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dataGridView_CellValidating);
            this.dataGridView.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.dataGridView_DataError);
            // 
            // SelectFileButton
            // 
            this.SelectFileButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SelectFileButton.Location = new System.Drawing.Point(162, 11);
            this.SelectFileButton.Name = "SelectFileButton";
            this.SelectFileButton.Size = new System.Drawing.Size(120, 40);
            this.SelectFileButton.TabIndex = 1;
            this.SelectFileButton.Text = "Select File";
            this.SelectFileButton.UseVisualStyleBackColor = true;
            this.SelectFileButton.Click += new System.EventHandler(this.SelectFileButton_Click);
            // 
            // ImportFileButton
            // 
            this.ImportFileButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ImportFileButton.Location = new System.Drawing.Point(606, 32);
            this.ImportFileButton.Name = "ImportFileButton";
            this.ImportFileButton.Size = new System.Drawing.Size(120, 40);
            this.ImportFileButton.TabIndex = 2;
            this.ImportFileButton.Text = "Import File";
            this.ImportFileButton.UseVisualStyleBackColor = true;
            this.ImportFileButton.Click += new System.EventHandler(this.ImportFileButton_Click);
            // 
            // ExportFileButton
            // 
            this.ExportFileButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ExportFileButton.Location = new System.Drawing.Point(1203, 32);
            this.ExportFileButton.Name = "ExportFileButton";
            this.ExportFileButton.Size = new System.Drawing.Size(120, 40);
            this.ExportFileButton.TabIndex = 3;
            this.ExportFileButton.Text = "Export File";
            this.ExportFileButton.UseVisualStyleBackColor = true;
            this.ExportFileButton.Click += new System.EventHandler(this.ExportFileButton_Click);
            // 
            // AddRowButton
            // 
            this.AddRowButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AddRowButton.Location = new System.Drawing.Point(39, 21);
            this.AddRowButton.Name = "AddRowButton";
            this.AddRowButton.Size = new System.Drawing.Size(120, 40);
            this.AddRowButton.TabIndex = 4;
            this.AddRowButton.Text = "Add Row";
            this.AddRowButton.UseVisualStyleBackColor = true;
            this.AddRowButton.Click += new System.EventHandler(this.AddRowButton_Click);
            // 
            // DeleteRowButton
            // 
            this.DeleteRowButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DeleteRowButton.Location = new System.Drawing.Point(166, 21);
            this.DeleteRowButton.Name = "DeleteRowButton";
            this.DeleteRowButton.Size = new System.Drawing.Size(120, 40);
            this.DeleteRowButton.TabIndex = 5;
            this.DeleteRowButton.Text = "Delete Row";
            this.DeleteRowButton.UseVisualStyleBackColor = true;
            this.DeleteRowButton.Click += new System.EventHandler(this.DeleteRowButton_Click);
            // 
            // CheckConflictButton
            // 
            this.CheckConflictButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CheckConflictButton.Location = new System.Drawing.Point(811, 32);
            this.CheckConflictButton.Name = "CheckConflictButton";
            this.CheckConflictButton.Size = new System.Drawing.Size(120, 40);
            this.CheckConflictButton.TabIndex = 7;
            this.CheckConflictButton.Text = "Check Conflict";
            this.CheckConflictButton.UseVisualStyleBackColor = true;
            this.CheckConflictButton.Click += new System.EventHandler(this.CheckConflictButton_Click);
            // 
            // excelFilePathTB
            // 
            this.excelFilePathTB.Location = new System.Drawing.Point(288, 22);
            this.excelFilePathTB.Name = "excelFilePathTB";
            this.excelFilePathTB.Size = new System.Drawing.Size(302, 20);
            this.excelFilePathTB.TabIndex = 8;
            // 
            // ExcelSheetNames
            // 
            this.ExcelSheetNames.BackColor = System.Drawing.SystemColors.Window;
            this.ExcelSheetNames.ForeColor = System.Drawing.SystemColors.ScrollBar;
            this.ExcelSheetNames.FormattingEnabled = true;
            this.ExcelSheetNames.Location = new System.Drawing.Point(288, 65);
            this.ExcelSheetNames.Name = "ExcelSheetNames";
            this.ExcelSheetNames.Size = new System.Drawing.Size(302, 21);
            this.ExcelSheetNames.TabIndex = 9;
            // 
            // SheetNamesLabel
            // 
            this.SheetNamesLabel.AutoSize = true;
            this.SheetNamesLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SheetNamesLabel.ForeColor = System.Drawing.Color.WhiteSmoke;
            this.SheetNamesLabel.Location = new System.Drawing.Point(163, 66);
            this.SheetNamesLabel.Name = "SheetNamesLabel";
            this.SheetNamesLabel.Size = new System.Drawing.Size(119, 17);
            this.SheetNamesLabel.TabIndex = 10;
            this.SheetNamesLabel.Text = "Sheet Names : ";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.SeaGreen;
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Controls.Add(this.GenerateReportButton);
            this.panel1.Controls.Add(this.ImportFileButton);
            this.panel1.Controls.Add(this.CheckConflictButton);
            this.panel1.Controls.Add(this.ExcelSheetNames);
            this.panel1.Controls.Add(this.SheetNamesLabel);
            this.panel1.Controls.Add(this.ExportFileButton);
            this.panel1.Controls.Add(this.SelectFileButton);
            this.panel1.Controls.Add(this.excelFilePathTB);
            this.panel1.Location = new System.Drawing.Point(12, 37);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1344, 103);
            this.panel1.TabIndex = 11;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.BackgroundImage")));
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pictureBox1.Location = new System.Drawing.Point(30, 6);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(92, 92);
            this.pictureBox1.TabIndex = 12;
            this.pictureBox1.TabStop = false;
            // 
            // GenerateReportButton
            // 
            this.GenerateReportButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GenerateReportButton.Location = new System.Drawing.Point(1047, 32);
            this.GenerateReportButton.Name = "GenerateReportButton";
            this.GenerateReportButton.Size = new System.Drawing.Size(138, 40);
            this.GenerateReportButton.TabIndex = 11;
            this.GenerateReportButton.Text = "Generate Reports";
            this.GenerateReportButton.UseVisualStyleBackColor = true;
            this.GenerateReportButton.Click += new System.EventHandler(this.GenerateReportButton_Click);
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.SeaGreen;
            this.panel2.Controls.Add(this.deleteEmptyColumns);
            this.panel2.Controls.Add(this.deleteEmptyRows);
            this.panel2.Controls.Add(this.DeleteRowButton);
            this.panel2.Controls.Add(this.AddRowButton);
            this.panel2.Location = new System.Drawing.Point(12, 649);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1344, 89);
            this.panel2.TabIndex = 12;
            // 
            // deleteEmptyColumns
            // 
            this.deleteEmptyColumns.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.deleteEmptyColumns.Location = new System.Drawing.Point(1144, 21);
            this.deleteEmptyColumns.Name = "deleteEmptyColumns";
            this.deleteEmptyColumns.Size = new System.Drawing.Size(179, 40);
            this.deleteEmptyColumns.TabIndex = 7;
            this.deleteEmptyColumns.Text = "Delete All Empty Columns";
            this.deleteEmptyColumns.UseVisualStyleBackColor = true;
            this.deleteEmptyColumns.Click += new System.EventHandler(this.deleteEmptyColumns_Click);
            // 
            // deleteEmptyRows
            // 
            this.deleteEmptyRows.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.deleteEmptyRows.Location = new System.Drawing.Point(970, 21);
            this.deleteEmptyRows.Name = "deleteEmptyRows";
            this.deleteEmptyRows.Size = new System.Drawing.Size(168, 40);
            this.deleteEmptyRows.TabIndex = 6;
            this.deleteEmptyRows.Text = "Delete All Empty Rows";
            this.deleteEmptyRows.UseVisualStyleBackColor = true;
            this.deleteEmptyRows.Click += new System.EventHandler(this.deleteEmptyRows_Click);
            // 
            // Images
            // 
            this.Images.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("Images.ImageStream")));
            this.Images.TransparentColor = System.Drawing.Color.Transparent;
            this.Images.Images.SetKeyName(0, "bearcat.jpg");
            // 
            // SchedulingAssistantForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.LightGray;
            this.ClientSize = new System.Drawing.Size(1362, 742);
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.Name = "SchedulingAssistantForm";
            this.Text = "SchedulingAssistantForm";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.Button SelectFileButton;
        private System.Windows.Forms.Button ImportFileButton;
        private System.Windows.Forms.Button ExportFileButton;
        private System.Windows.Forms.Button AddRowButton;
        private System.Windows.Forms.Button DeleteRowButton;
        private System.Windows.Forms.Button CheckConflictButton;
        private System.Windows.Forms.TextBox excelFilePathTB;
        private System.Windows.Forms.ComboBox ExcelSheetNames;
        private System.Windows.Forms.Label SheetNamesLabel;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button deleteEmptyColumns;
        private System.Windows.Forms.Button deleteEmptyRows;
        private System.Windows.Forms.Button GenerateReportButton;
        private System.Windows.Forms.ImageList Images;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}