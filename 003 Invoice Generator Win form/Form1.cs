using System;
using System.IO;
using System.Windows.Forms;
using GenerateInvoice;

namespace InvoiceGenerator_WinForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            selectOutput.Enabled = false;
            validate.Enabled = false;
            process.Enabled = false;
        }

        private void SelectInput_Click(object sender, EventArgs e)
        {
            selectInputFIle_Dialog.FileName = "";
            if (selectInputFIle_Dialog.ShowDialog() == DialogResult.OK)
            {
                inputFileAddress.Text = selectInputFIle_Dialog.FileName;
                progressText.Text = "Input Data File Selected";
                selectOutput.Enabled = true;
                progress.Value = 20;
            }
        }

        private void SelectOutput_Click(object sender, EventArgs e)
        {
            if (File.Exists(inputFileAddress.Text))
            {
                selectOutputPath_Dialog.SelectedPath = Path.GetDirectoryName(inputFileAddress.Text);
                if (selectOutputPath_Dialog.ShowDialog() == DialogResult.OK)
                {
                    outputPathAddress.Text = selectOutputPath_Dialog.SelectedPath;
                    progressText.Text = "Output Path Selected";
                    validate.Enabled = true;
                    process.Enabled = true;
                    progress.Value = 40;
                }
            }
        }

        private void Validate_Click(object sender, EventArgs e)
        {
            Entry.MainEntry(selectInputFIle_Dialog.FileName, selectOutputPath_Dialog.SelectedPath, progressText, progress);
            progress.Value = 50;
        }

        private void Process_Click(object sender, EventArgs e)
        {
            progress.Value =60;
            Entry.MainEntry(selectInputFIle_Dialog.FileName, selectOutputPath_Dialog.SelectedPath, progressText, progress, true);
            selectOutput.Enabled = false;
            validate.Enabled = false;
            process.Enabled = false;
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void progress_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
