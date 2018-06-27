namespace ExportToSpreadsheet
{
    partial class Form1
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
            this.btnExort = new System.Windows.Forms.Button();
            this.progBar = new System.Windows.Forms.ProgressBar();
            this.txtState = new System.Windows.Forms.TextBox();
            this.lblState = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnExort
            // 
            this.btnExort.Location = new System.Drawing.Point(12, 12);
            this.btnExort.Name = "btnExort";
            this.btnExort.Size = new System.Drawing.Size(75, 23);
            this.btnExort.TabIndex = 0;
            this.btnExort.Text = "Export";
            this.btnExort.UseVisualStyleBackColor = true;
            this.btnExort.Click += new System.EventHandler(this.btnExort_Click);
            // 
            // progBar
            // 
            this.progBar.Location = new System.Drawing.Point(12, 94);
            this.progBar.Name = "progBar";
            this.progBar.Size = new System.Drawing.Size(560, 23);
            this.progBar.TabIndex = 2;
            // 
            // txtState
            // 
            this.txtState.Enabled = false;
            this.txtState.Location = new System.Drawing.Point(93, 14);
            this.txtState.Multiline = true;
            this.txtState.Name = "txtState";
            this.txtState.Size = new System.Drawing.Size(479, 74);
            this.txtState.TabIndex = 3;
            // 
            // lblState
            // 
            this.lblState.AutoSize = true;
            this.lblState.Location = new System.Drawing.Point(180, 120);
            this.lblState.Name = "lblState";
            this.lblState.Size = new System.Drawing.Size(0, 13);
            this.lblState.TabIndex = 4;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 145);
            this.Controls.Add(this.lblState);
            this.Controls.Add(this.txtState);
            this.Controls.Add(this.progBar);
            this.Controls.Add(this.btnExort);
            this.Name = "Form1";
            this.Text = "ZDHC Exporter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnExort;
        private System.Windows.Forms.ProgressBar progBar;
        private System.Windows.Forms.TextBox txtState;
        private System.Windows.Forms.Label lblState;
    }
}

