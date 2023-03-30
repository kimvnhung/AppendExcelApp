using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Append_Excel
{
    public partial class Form1 : Form
    {
        private string mLastOpenFolder = @"E:\Freelancer\firstjob";
        private List<string> mSelectedFile = new List<string>();

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
    }
}
