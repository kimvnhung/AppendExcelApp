using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Append_Excel
{
    public partial class Form1 : Form
    {
        private string mLastOpenFolder = @"E:\Freelancer\firstjob";
        private List<string> mSelectedFile = new List<string>();
        private bool mIsProcessing = false;

        public Form1()
        {
            InitializeComponent();
            selectedFileList.HeaderStyle = ColumnHeaderStyle.None;
            selectedFileList.View = View.Details;
        }

        private void openFileBtn_Click(object sender, EventArgs e)
        {
            var openFile = new OpenFileDialog
            {
                InitialDirectory = mLastOpenFolder,
                Filter = "Excel Files(*.xls;*.xlsx;*.csv)|*.xls;*.xlsx;*.csv;|All Files(*;)|*;"
            };
            openFile.Title = "Open Excel or CSV Files";
            openFile.Multiselect = true;

            if (openFile.ShowDialog() == DialogResult.OK) {
                mLastOpenFolder = System.IO.Path.GetDirectoryName(openFile.FileName);
                mSelectedFile.Clear();
                selectedFileList.Items.Clear();
               foreach (String fileName in openFile.FileNames)
                {
                    mSelectedFile.Add(fileName);
                    selectedFileList.Items.Add(System.IO.Path.GetFileName(fileName));
                }
            }
        }

        private void appendBtn_Click(object sender, EventArgs e)
        {
            if (mSelectedFile.Count == 0) {
                MessageBox.Show("Has no file to merge!","Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                InitialDirectory = mLastOpenFolder,
                FileName = System.IO.Path.GetFileName(mLastOpenFolder),
                Filter = "Excel Files (*.xlsx;)|*.xlsx;|CSV Files(*.csv)|*.csv",
                Title = "Save Files",
                
            };

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = saveDialog.FileName;
                Task.Run(startProcessing);
            }
        }

        private async Task startProcessing()
        {
            mIsProcessing = true;
            updateView();
            Thread.Sleep(2000);
            mIsProcessing = false;
            updateView();
        }


        private void updateView() 
        {
            if (mIsProcessing)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    appendBtn.Text = "Cancel";
                });
            }
            else
            {
                this.Invoke((MethodInvoker)delegate
                {
                    appendBtn.Text = "Append";
                });
            }
        }
    }
}
