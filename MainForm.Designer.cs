namespace QueryToExcel
{
    partial class MainForm
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
            this.queryTextBox = new System.Windows.Forms.TextBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.exportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.connectionDropdown = new System.Windows.Forms.ComboBox();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.statusCurrentOperation = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusElapsed = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusRowCount = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusFileSize = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusTimer = new System.Windows.Forms.Timer(this.components);
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // queryTextBox
            // 
            this.queryTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.queryTextBox.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.queryTextBox.Location = new System.Drawing.Point(0, 80);
            this.queryTextBox.Multiline = true;
            this.queryTextBox.Name = "queryTextBox";
            this.queryTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.queryTextBox.Size = new System.Drawing.Size(842, 399);
            this.queryTextBox.TabIndex = 0;
            this.queryTextBox.WordWrap = false;
            this.queryTextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.queryTextBox_KeyDown);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exportToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(842, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // exportToolStripMenuItem
            // 
            this.exportToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.excelToolStripMenuItem});
            this.exportToolStripMenuItem.Name = "exportToolStripMenuItem";
            this.exportToolStripMenuItem.Size = new System.Drawing.Size(52, 20);
            this.exportToolStripMenuItem.Text = "E&xport";
            // 
            // excelToolStripMenuItem
            // 
            this.excelToolStripMenuItem.Name = "excelToolStripMenuItem";
            this.excelToolStripMenuItem.Size = new System.Drawing.Size(117, 22);
            this.excelToolStripMenuItem.Text = "To &Excel";
            this.excelToolStripMenuItem.Click += new System.EventHandler(this.excelToolStripMenuItem_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.connectionDropdown);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(0, 24);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(8);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(842, 56);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Connection";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "&Connection";
            // 
            // connectionDropdown
            // 
            this.connectionDropdown.FormattingEnabled = true;
            this.connectionDropdown.Location = new System.Drawing.Point(79, 19);
            this.connectionDropdown.Name = "connectionDropdown";
            this.connectionDropdown.Size = new System.Drawing.Size(192, 21);
            this.connectionDropdown.TabIndex = 0;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statusCurrentOperation,
            this.statusElapsed,
            this.statusRowCount,
            this.statusFileSize});
            this.statusStrip1.Location = new System.Drawing.Point(0, 479);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(842, 22);
            this.statusStrip1.TabIndex = 3;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // statusCurrentOperation
            // 
            this.statusCurrentOperation.Name = "statusCurrentOperation";
            this.statusCurrentOperation.Size = new System.Drawing.Size(635, 17);
            this.statusCurrentOperation.Spring = true;
            this.statusCurrentOperation.Text = "Ready";
            this.statusCurrentOperation.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // statusElapsed
            // 
            this.statusElapsed.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter;
            this.statusElapsed.Name = "statusElapsed";
            this.statusElapsed.Size = new System.Drawing.Size(89, 17);
            this.statusElapsed.Text = "Elapsed: 0:00:00";
            // 
            // statusRowCount
            // 
            this.statusRowCount.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter;
            this.statusRowCount.Name = "statusRowCount";
            this.statusRowCount.Size = new System.Drawing.Size(47, 17);
            this.statusRowCount.Text = "Rows: 0";
            // 
            // statusFileSize
            // 
            this.statusFileSize.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter;
            this.statusFileSize.Name = "statusFileSize";
            this.statusFileSize.Size = new System.Drawing.Size(56, 17);
            this.statusFileSize.Text = "Size: 0 KB";
            // 
            // statusTimer
            // 
            this.statusTimer.Enabled = true;
            this.statusTimer.Interval = 500;
            this.statusTimer.Tick += new System.EventHandler(this.statusTimer_Tick);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(842, 501);
            this.Controls.Add(this.queryTextBox);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainForm";
            this.Text = "Query To Excel";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox queryTextBox;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem exportToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excelToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox connectionDropdown;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel statusCurrentOperation;
        private System.Windows.Forms.ToolStripStatusLabel statusElapsed;
        private System.Windows.Forms.ToolStripStatusLabel statusRowCount;
        private System.Windows.Forms.ToolStripStatusLabel statusFileSize;
        private System.Windows.Forms.Timer statusTimer;
    }
}

