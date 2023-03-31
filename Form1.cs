using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
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
        private AppendHandler mHandler;

        public Form1()
        {
            InitializeComponent();
            selectedFileList.HeaderStyle = ColumnHeaderStyle.None;
            selectedFileList.View = View.Details;

            mHandler = new AppendHandler();
            mHandler.StatusChanged += MHandler_StatusChanged;
            mHandler.ShowMessage += MHandler_ShowMessage;
        }

        private void MHandler_ShowMessage(object sender, string e)
        {
            this.Invoke((MethodInvoker)delegate
            {
                appendStatusLb.Text = mHandler.Message;
            });
        }

        private void MHandler_StatusChanged(object sender, EventArgs e)
        {
            if (mHandler.IsProcessing)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    openFileBtn.Enabled = false;
                    appendBtn.Text = "Cancel";
                    if (mHandler.PercentageProcess >= 0 && mHandler.PercentageProcess <= 100)
                    {
                        progressBar1.Value = mHandler.PercentageProcess;
                    }
                    processStatusLb.Text = ConvertExecuted(mHandler.ExecutedTime)+" / "+ConvertEstimate(mHandler.EstimatedTime);
                });

            }
            else
            {
                this.Invoke((MethodInvoker)delegate
                {
                    appendBtn.Text = "Append";
                    openFileBtn.Enabled = true;
                    progressBar1.Value = 0;
                });
            }
        }



        public static string ConvertExecuted(long milliseconds)
        {
            TimeSpan time = TimeSpan.FromMilliseconds(milliseconds);
            return string.Format("{0:D2}:{1:D2}:{2:D2}:{3:D2}ms",
                (int)time.TotalHours,
                time.Minutes,
                time.Seconds,
                time.Milliseconds);
        }

        public static string ConvertEstimate(long milliseconds) 
        {
            TimeSpan time = TimeSpan.FromMilliseconds(milliseconds);
            //return string.Format("{0:D2}h{1:D2}m{2:D2}s",
            //    (int)time.TotalHours,
            //    time.Minutes,
            //    time.Seconds);
            if(time.Hours > 0)
            {
                return string.Format("~ {0:D2}h{1:D2}m",
                    (int)time.TotalHours,
                    time.Minutes);
            }else if(time.Minutes > 0)
            {
                return string.Format("~ {0:D2}m{1:D2}s",
                    (int)time.TotalMinutes,
                    time.Seconds);
            }else if(time.TotalSeconds > 2)
            {
                return string.Format("~ {0:D2}s",
                    (int)time.TotalSeconds+1);
            }else
            {
                return string.Format("< {0:D2}s",
                   (int)time.TotalSeconds + 1);
            }
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

        private async void appendBtn_Click(object sender, EventArgs e)
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

                _ = Task.Run(async () => { await mHandler.TimeEstimateHandler(); });
                _ = mHandler.StartProcessing(mSelectedFile,fileName);
            }
        }

        private void cleanProcessBtn_Click(object sender, EventArgs e)
        {
            //// Get all running instances of Excel
            //Process[] processes = Process.GetProcessesByName("Excel");

            //// Close each instance of Excel
            //foreach (Process process in processes)
            //{
            //    // Try to close the Excel process gracefully
            //    try
            //    {
            //        // Get the Excel Application object
            //        Microsoft.Office.Interop.Excel.Application excelApp = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

            //        // Close all open workbooks
            //        excelApp.Workbooks.Close();

            //        // Quit the Excel Application object
            //        excelApp.Quit();
            //    }
            //    catch
            //    {
            //        // Ignore any exceptions and kill the process forcibly
            //    }

            //    // Kill the Excel process forcibly
            //    process.Kill();
            //}
            // Start a new process to run the taskkill command
            Process process = new Process();
            process.StartInfo.FileName = "taskkill";
            process.StartInfo.Arguments = "/f /im excel.exe";
            process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            process.Start();
            process.WaitForExit();
        }
    }
}
