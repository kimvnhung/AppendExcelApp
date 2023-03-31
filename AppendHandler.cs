using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Append_Excel
{
    class AppendHandler
    {
        public event EventHandler StatusChanged;

        private bool mIsProcessing = false;
        private int mPercentageProcess = 0;
        private int mExecutedTime = 0;
        private int mEstimateTime = 0;

        public bool IsProcessing { 
            private set { 
                if(value != mIsProcessing)
                {
                    mIsProcessing = value;
                    OnStatusChanged();
                } 
            }
            get { return mIsProcessing; } 
        }
        public int PercentageProcess { 
            private set
            {
                if(value != mPercentageProcess)
                {
                    mPercentageProcess = value;
                    OnStatusChanged();
                }
            }
            get { return mPercentageProcess; } 
        }

        public int ExceutedTime
        {
            private set
            {
                if (value != mExecutedTime)
                {
                    mExecutedTime = value;
                    OnStatusChanged() ;
                }
            }
            get { return mExecutedTime; }
        }

        public int EstimatedTime { 
            private set
            {
                if( value != mEstimateTime)
                {
                    mEstimateTime = value;
                    OnStatusChanged();
                }
            }
            get { return mEstimateTime; } 
        }
        public AppendHandler() { }


        public async Task StartProcessing(List<string> selectedFiles)
        {
            _ = TimeEstimateHandler();
            IsProcessing = true;
            
            IsProcessing = false;

        }

        public async Task TimeEstimateHandler()
        {
            DateTime start = DateTime.Now;
            while (PercentageProcess >= 0 && PercentageProcess < 100)
            {
                await Task.Delay(100);

                TimeSpan elapsed = DateTime.Now - start;

                ExceutedTime = (int)elapsed.TotalMilliseconds;
                EstimatedTime = ExceutedTime * 100 / PercentageProcess;
            }
        }

        private async Task<List<Worksheet>> OpenDataFiles(List<string> selectedFiles)
        {
            await Task.Delay(100);
            return new List<Worksheet>();
        }

        private async Task<bool> Appending(List<Worksheet> workSheets)
        {
            return false;
        }

        private async Task<bool> SaveFile(string filePath)
        {
            return false;
        }

        protected virtual void OnStatusChanged() {
            StatusChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
