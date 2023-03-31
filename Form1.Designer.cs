namespace Append_Excel
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
            this.label1 = new System.Windows.Forms.Label();
            this.selectedFileList = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.openFileBtn = new System.Windows.Forms.Button();
            this.appendBtn = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.processStatusLb = new System.Windows.Forms.Label();
            this.cleanProcessBtn = new System.Windows.Forms.Button();
            this.appendStatusLb = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Selected Files";
            // 
            // selectedFileList
            // 
            this.selectedFileList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
            this.selectedFileList.HideSelection = false;
            this.selectedFileList.Location = new System.Drawing.Point(15, 25);
            this.selectedFileList.Name = "selectedFileList";
            this.selectedFileList.Size = new System.Drawing.Size(252, 413);
            this.selectedFileList.TabIndex = 1;
            this.selectedFileList.UseCompatibleStateImageBehavior = false;
            this.selectedFileList.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Width = 256;
            // 
            // openFileBtn
            // 
            this.openFileBtn.Location = new System.Drawing.Point(273, 25);
            this.openFileBtn.Name = "openFileBtn";
            this.openFileBtn.Size = new System.Drawing.Size(75, 23);
            this.openFileBtn.TabIndex = 2;
            this.openFileBtn.Text = "Open Files";
            this.openFileBtn.UseVisualStyleBackColor = true;
            this.openFileBtn.Click += new System.EventHandler(this.openFileBtn_Click);
            // 
            // appendBtn
            // 
            this.appendBtn.Location = new System.Drawing.Point(354, 25);
            this.appendBtn.Name = "appendBtn";
            this.appendBtn.Size = new System.Drawing.Size(75, 23);
            this.appendBtn.TabIndex = 3;
            this.appendBtn.Text = "Append";
            this.appendBtn.UseVisualStyleBackColor = true;
            this.appendBtn.Click += new System.EventHandler(this.appendBtn_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(354, 238);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(387, 23);
            this.progressBar1.TabIndex = 4;
            // 
            // processStatusLb
            // 
            this.processStatusLb.AutoSize = true;
            this.processStatusLb.Location = new System.Drawing.Point(351, 222);
            this.processStatusLb.Name = "processStatusLb";
            this.processStatusLb.Size = new System.Drawing.Size(37, 13);
            this.processStatusLb.TabIndex = 5;
            this.processStatusLb.Text = "Status";
            // 
            // cleanProcessBtn
            // 
            this.cleanProcessBtn.Location = new System.Drawing.Point(713, 24);
            this.cleanProcessBtn.Name = "cleanProcessBtn";
            this.cleanProcessBtn.Size = new System.Drawing.Size(75, 23);
            this.cleanProcessBtn.TabIndex = 6;
            this.cleanProcessBtn.Text = "Clean Process";
            this.cleanProcessBtn.UseVisualStyleBackColor = true;
            this.cleanProcessBtn.Click += new System.EventHandler(this.cleanProcessBtn_Click);
            // 
            // appendStatusLb
            // 
            this.appendStatusLb.AutoSize = true;
            this.appendStatusLb.Location = new System.Drawing.Point(351, 264);
            this.appendStatusLb.MaximumSize = new System.Drawing.Size(387, 0);
            this.appendStatusLb.Name = "appendStatusLb";
            this.appendStatusLb.Size = new System.Drawing.Size(0, 13);
            this.appendStatusLb.TabIndex = 8;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.appendStatusLb);
            this.Controls.Add(this.cleanProcessBtn);
            this.Controls.Add(this.processStatusLb);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.appendBtn);
            this.Controls.Add(this.openFileBtn);
            this.Controls.Add(this.selectedFileList);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Append Excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListView selectedFileList;
        private System.Windows.Forms.Button openFileBtn;
        private System.Windows.Forms.Button appendBtn;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label processStatusLb;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.Button cleanProcessBtn;
        private System.Windows.Forms.Label appendStatusLb;
    }
}

