namespace InvoiceGenerator_WinForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.selectInput = new System.Windows.Forms.Button();
            this.selectOutput = new System.Windows.Forms.Button();
            this.inputFileAddress = new System.Windows.Forms.TextBox();
            this.outputPathAddress = new System.Windows.Forms.TextBox();
            this.validate = new System.Windows.Forms.Button();
            this.process = new System.Windows.Forms.Button();
            this.progress = new System.Windows.Forms.ProgressBar();
            this.exit = new System.Windows.Forms.Button();
            this.selectInputFIle_Dialog = new System.Windows.Forms.OpenFileDialog();
            this.selectOutputPath_Dialog = new System.Windows.Forms.FolderBrowserDialog();
            this.progressText = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // selectInput
            // 
            this.selectInput.Location = new System.Drawing.Point(45, 40);
            this.selectInput.Margin = new System.Windows.Forms.Padding(2);
            this.selectInput.Name = "selectInput";
            this.selectInput.Size = new System.Drawing.Size(135, 31);
            this.selectInput.TabIndex = 0;
            this.selectInput.Text = "Select Input File";
            this.selectInput.UseVisualStyleBackColor = true;
            this.selectInput.Click += new System.EventHandler(this.SelectInput_Click);
            // 
            // selectOutput
            // 
            this.selectOutput.Location = new System.Drawing.Point(45, 105);
            this.selectOutput.Margin = new System.Windows.Forms.Padding(2);
            this.selectOutput.Name = "selectOutput";
            this.selectOutput.Size = new System.Drawing.Size(135, 34);
            this.selectOutput.TabIndex = 1;
            this.selectOutput.Text = "Select Output Path";
            this.selectOutput.UseVisualStyleBackColor = true;
            this.selectOutput.Click += new System.EventHandler(this.SelectOutput_Click);
            // 
            // inputFileAddress
            // 
            this.inputFileAddress.Location = new System.Drawing.Point(232, 40);
            this.inputFileAddress.Margin = new System.Windows.Forms.Padding(2);
            this.inputFileAddress.Multiline = true;
            this.inputFileAddress.Name = "inputFileAddress";
            this.inputFileAddress.Size = new System.Drawing.Size(328, 31);
            this.inputFileAddress.TabIndex = 2;
            // 
            // outputPathAddress
            // 
            this.outputPathAddress.Location = new System.Drawing.Point(232, 105);
            this.outputPathAddress.Margin = new System.Windows.Forms.Padding(2);
            this.outputPathAddress.Multiline = true;
            this.outputPathAddress.Name = "outputPathAddress";
            this.outputPathAddress.Size = new System.Drawing.Size(328, 34);
            this.outputPathAddress.TabIndex = 3;
            // 
            // validate
            // 
            this.validate.Location = new System.Drawing.Point(122, 240);
            this.validate.Margin = new System.Windows.Forms.Padding(2);
            this.validate.Name = "validate";
            this.validate.Size = new System.Drawing.Size(108, 30);
            this.validate.TabIndex = 4;
            this.validate.Text = "Validate Excel";
            this.validate.UseVisualStyleBackColor = true;
            this.validate.Click += new System.EventHandler(this.Validate_Click);
            // 
            // process
            // 
            this.process.Location = new System.Drawing.Point(392, 240);
            this.process.Margin = new System.Windows.Forms.Padding(2);
            this.process.Name = "process";
            this.process.Size = new System.Drawing.Size(105, 30);
            this.process.TabIndex = 5;
            this.process.Text = "Convert";
            this.process.UseVisualStyleBackColor = true;
            this.process.Click += new System.EventHandler(this.Process_Click);
            // 
            // progress
            // 
            this.progress.Location = new System.Drawing.Point(122, 181);
            this.progress.Margin = new System.Windows.Forms.Padding(2);
            this.progress.MarqueeAnimationSpeed = 1000;
            this.progress.Maximum = 1000;
            this.progress.Name = "progress";
            this.progress.Size = new System.Drawing.Size(375, 24);
            this.progress.TabIndex = 6;
            this.progress.Click += new System.EventHandler(this.progress_Click);
            // 
            // exit
            // 
            this.exit.Location = new System.Drawing.Point(9, 327);
            this.exit.Margin = new System.Windows.Forms.Padding(2);
            this.exit.Name = "exit";
            this.exit.Size = new System.Drawing.Size(84, 28);
            this.exit.TabIndex = 7;
            this.exit.Text = "Exit";
            this.exit.UseVisualStyleBackColor = true;
            this.exit.Click += new System.EventHandler(this.Exit_Click);
            // 
            // selectInputFIle_Dialog
            // 
            this.selectInputFIle_Dialog.FileName = "openFileDialog1";
            // 
            // progressText
            // 
            this.progressText.AutoSize = true;
            this.progressText.Location = new System.Drawing.Point(130, 207);
            this.progressText.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.progressText.Name = "progressText";
            this.progressText.Size = new System.Drawing.Size(0, 13);
            this.progressText.TabIndex = 8;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.BackgroundImage")));
            this.pictureBox1.Location = new System.Drawing.Point(563, 295);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(63, 59);
            this.pictureBox1.TabIndex = 9;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(661, 366);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.progressText);
            this.Controls.Add(this.exit);
            this.Controls.Add(this.progress);
            this.Controls.Add(this.process);
            this.Controls.Add(this.validate);
            this.Controls.Add(this.outputPathAddress);
            this.Controls.Add(this.inputFileAddress);
            this.Controls.Add(this.selectOutput);
            this.Controls.Add(this.selectInput);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.Text = "Excel to Pdf Invoice";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button selectInput;
        private System.Windows.Forms.Button selectOutput;
        private System.Windows.Forms.TextBox inputFileAddress;
        private System.Windows.Forms.TextBox outputPathAddress;
        private System.Windows.Forms.Button validate;
        private System.Windows.Forms.Button process;
        private System.Windows.Forms.ProgressBar progress;
        private System.Windows.Forms.Button exit;
        private System.Windows.Forms.OpenFileDialog selectInputFIle_Dialog;
        private System.Windows.Forms.FolderBrowserDialog selectOutputPath_Dialog;
        private System.Windows.Forms.Label progressText;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}

