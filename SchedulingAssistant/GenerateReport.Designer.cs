namespace SchedulingAssistant
{
    partial class GenerateReport
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
            this.label1 = new System.Windows.Forms.Label();
            this.generateReportByCB = new System.Windows.Forms.ComboBox();
            this.reportOfCB = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.GeneratePDFButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 68);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(95, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Generate report by";
            // 
            // generateReportByCB
            // 
            this.generateReportByCB.FormattingEnabled = true;
            this.generateReportByCB.Location = new System.Drawing.Point(123, 65);
            this.generateReportByCB.Name = "generateReportByCB";
            this.generateReportByCB.Size = new System.Drawing.Size(305, 21);
            this.generateReportByCB.TabIndex = 2;
            this.generateReportByCB.SelectedIndexChanged += new System.EventHandler(this.generateReportByCB_SelectedIndexChanged);
            this.generateReportByCB.SelectionChangeCommitted += new System.EventHandler(this.generateReportByCB_SelectionChangeCommitted);
            // 
            // reportOfCB
            // 
            this.reportOfCB.FormattingEnabled = true;
            this.reportOfCB.Location = new System.Drawing.Point(59, 142);
            this.reportOfCB.Name = "reportOfCB";
            this.reportOfCB.Size = new System.Drawing.Size(344, 21);
            this.reportOfCB.TabIndex = 3;
            this.reportOfCB.SelectedIndexChanged += new System.EventHandler(this.reportOfCB_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(151, 116);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(141, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Select from below dropdown";
            // 
            // GeneratePDFButton
            // 
            this.GeneratePDFButton.Location = new System.Drawing.Point(15, 280);
            this.GeneratePDFButton.Name = "GeneratePDFButton";
            this.GeneratePDFButton.Size = new System.Drawing.Size(428, 36);
            this.GeneratePDFButton.TabIndex = 5;
            this.GeneratePDFButton.Text = "Generate PDF";
            this.GeneratePDFButton.UseVisualStyleBackColor = true;
            this.GeneratePDFButton.Click += new System.EventHandler(this.GeneratePDFButton_Click);
            // 
            // GenerateReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(455, 337);
            this.Controls.Add(this.GeneratePDFButton);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.reportOfCB);
            this.Controls.Add(this.generateReportByCB);
            this.Controls.Add(this.label1);
            this.Name = "GenerateReport";
            this.Text = "GenerateReport";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox generateReportByCB;
        private System.Windows.Forms.ComboBox reportOfCB;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button GeneratePDFButton;
    }
}