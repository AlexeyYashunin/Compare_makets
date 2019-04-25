namespace Compare_makets {
    partial class Form1 {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if(disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.labelPBar = new System.Windows.Forms.Label();
            this.btnCompare = new System.Windows.Forms.Button();
            this.pbSecondFile = new System.Windows.Forms.PictureBox();
            this.pbFirstFile = new System.Windows.Forms.PictureBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.pbSecondFile)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbFirstFile)).BeginInit();
            this.SuspendLayout();
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            // 
            // progressBar1
            // 
            this.progressBar1.ForeColor = System.Drawing.Color.Black;
            this.progressBar1.Location = new System.Drawing.Point(12, 238);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(571, 23);
            this.progressBar1.TabIndex = 5;
            // 
            // labelPBar
            // 
            this.labelPBar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.labelPBar.AutoSize = true;
            this.labelPBar.BackColor = System.Drawing.Color.Transparent;
            this.labelPBar.Font = new System.Drawing.Font("Century Gothic", 10F, System.Drawing.FontStyle.Bold);
            this.labelPBar.Location = new System.Drawing.Point(279, 269);
            this.labelPBar.Name = "labelPBar";
            this.labelPBar.Size = new System.Drawing.Size(32, 17);
            this.labelPBar.TabIndex = 6;
            this.labelPBar.Text = "0 %";
            // 
            // btnCompare
            // 
            this.btnCompare.BackgroundImage = global::Compare_makets.Properties.Resources._source_compare_2;
            this.btnCompare.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnCompare.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCompare.Font = new System.Drawing.Font("Century Gothic", 12F);
            this.btnCompare.Location = new System.Drawing.Point(241, 72);
            this.btnCompare.Name = "btnCompare";
            this.btnCompare.Size = new System.Drawing.Size(110, 87);
            this.btnCompare.TabIndex = 2;
            this.btnCompare.UseVisualStyleBackColor = true;
            this.btnCompare.Click += new System.EventHandler(this.btnCompare_Click);
            // 
            // pbSecondFile
            // 
            this.pbSecondFile.Enabled = false;
            this.pbSecondFile.Image = global::Compare_makets.Properties.Resources._source_FGIS;
            this.pbSecondFile.InitialImage = null;
            this.pbSecondFile.Location = new System.Drawing.Point(356, 12);
            this.pbSecondFile.Name = "pbSecondFile";
            this.pbSecondFile.Size = new System.Drawing.Size(227, 200);
            this.pbSecondFile.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pbSecondFile.TabIndex = 1;
            this.pbSecondFile.TabStop = false;
            // 
            // pbFirstFile
            // 
            this.pbFirstFile.Enabled = false;
            this.pbFirstFile.Image = global::Compare_makets.Properties.Resources._source_Certsys;
            this.pbFirstFile.InitialImage = null;
            this.pbFirstFile.Location = new System.Drawing.Point(12, 12);
            this.pbFirstFile.Name = "pbFirstFile";
            this.pbFirstFile.Size = new System.Drawing.Size(223, 200);
            this.pbFirstFile.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pbFirstFile.TabIndex = 0;
            this.pbFirstFile.TabStop = false;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog2";
            // 
            // Form1
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(244)))), ((int)(((byte)(244)))), ((int)(((byte)(244)))));
            this.ClientSize = new System.Drawing.Size(595, 295);
            this.Controls.Add(this.labelPBar);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnCompare);
            this.Controls.Add(this.pbSecondFile);
            this.Controls.Add(this.pbFirstFile);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Compare makets";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.Form1_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.Form1_DragEnter);
            this.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.Form1_MouseDoubleClick);
            ((System.ComponentModel.ISupportInitialize)(this.pbSecondFile)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pbFirstFile)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pbFirstFile;
        private System.Windows.Forms.PictureBox pbSecondFile;
        private System.Windows.Forms.Button btnCompare;
        public System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label labelPBar;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
    }
}

