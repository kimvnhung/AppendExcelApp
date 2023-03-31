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
        public AppendHandler() { }


        public async Task StartProcessing()
        {
            PercentageProcess = 0;
            IsProcessing = true;
            await Task.Delay(1000);
            PercentageProcess = 20;
            await Task.Delay(1000);
            PercentageProcess = 80;
            await Task.Delay(1000);
            PercentageProcess = 100;
            IsProcessing = false;

        }

        protected virtual void OnStatusChanged() {
            StatusChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
