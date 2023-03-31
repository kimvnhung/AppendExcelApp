using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Append_Excel
{
    class AppendHandler
    {
        public event EventHandler StatusChanged;
        public event EventHandler<string> ShowMessage;
        private static object mLocker = new object();

        private bool mIsProcessing = false;
        private int mPercentageProcess = 0;
        private int mExecutedTime = 0;
        private int mEstimateTime = 0;
        private Application mExcel = null;

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

        private string mMessage = "";
        public string Message
        {
            private set
            {
                if(value != mMessage)
                {
                    mMessage = value;
                    OnShowMessage();
                }
            }
            get { return mMessage; }
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


        public async Task StartProcessing(List<string> selectedFiles,string savePath)
        {
            if(selectedFiles.Count == 0)
            {
                Message = "Has no file selected";
                return;
            }
            PercentageProcess = 0;
            TimeEstimateHandler();
            IsProcessing = true;
            //handling
            try
            {
                mExcel = new Application();
                Workbook wbResult = mExcel.Workbooks.Add();
                Workbook wbHandle = mExcel.Workbooks.Add();
                await OpenDataFiles(selectedFiles, wbHandle);
                if (await Appending(wbHandle))
                {
                    await SaveFile(wbResult, wbHandle, savePath);
                    Message = "Save file to " + savePath;
                    wbResult.Close(true);
                    wbHandle.Close(false);
                    mExcel.Quit();
                    return;
                }
                wbResult.Close(false);
                wbHandle.Close(false);
                mExcel.Quit();
            }catch(Exception ex)
            {
                Message=ex.Message;
                IsProcessing = false;
                return;
            }
            IsProcessing = false;
            Message = "Append Failed!";
        }

        public async Task TimeEstimateHandler()
        {
            DateTime start = DateTime.Now;
            while (IsProcessing && PercentageProcess >= 0 && PercentageProcess < 100)
            {
                Console.WriteLine(PercentageProcess);
                await Task.Delay(10);

                TimeSpan elapsed = DateTime.Now - start;

                ExceutedTime = (int)elapsed.TotalMilliseconds;
                EstimatedTime = ExceutedTime * 100 / PercentageProcess;
            }
            PercentageProcess = 0;
        }

        private async Task<bool> OpenDataFiles(List<string> selectedFiles, Workbook wbHandle)
        {
            int deltaPercentage = 30/selectedFiles.Count;
            foreach (string file in selectedFiles)
            {
                var fileExt = System.IO.Path.GetExtension(file);
                if (fileExt == ".xlsx" || fileExt == ".xls")
                {
                    try
                    {
                        await OpenXLSXFiles(file, wbHandle);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        continue;
                    }
                }
                else if(fileExt == ".csv")
                {
                    try
                    {
                        await OpenCSVFiles(file, wbHandle);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        continue;
                    }
                }
                PercentageProcess += deltaPercentage;
            }
            return true;
        }

        private async Task<bool> OpenXLSXFiles(string xlsxFiles, Workbook wbHandle)
        {
            Workbook workbook = mExcel.Workbooks.Open(xlsxFiles);
            Console.WriteLine(xlsxFiles+" WorkSheet Count : "+workbook.Worksheets.Count);
            foreach(Worksheet worksheet in workbook.Worksheets)
            {
                try
                {
                    
                    //PrintData(worksheet);
                    // Add a new sheet to the destination workbook
                    Worksheet destinationSheet = wbHandle.Worksheets.Add();

                    // Set the destination sheet name to match the source sheet name
                    destinationSheet.Name = worksheet.Name+wbHandle.Worksheets.Count;

                    // Copy the entire source sheet to the destination sheet
                    Copy(worksheet, destinationSheet);
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    workbook.Close(false);
                    return false;
                }
            }
            workbook.Close(false);
            return true;
        }

        private void Copy(Worksheet sourceSheet,Worksheet destinationSheet)
        {
            // Get the range of cells in the source worksheet
            Range sourceRange = sourceSheet.UsedRange;

            // Loop through each cell in the source range and copy its value and formatting to the destination worksheet
            for (int row = 1; row <= sourceRange.Rows.Count; row++)
            {
                for (int column = 1; column <= sourceRange.Columns.Count; column++)
                {
                    // Get the cell in the source range and the corresponding cell in the destination worksheet
                    Range sourceCell = sourceRange.Cells[row, column];
                    Range destinationCell = destinationSheet.Cells[row, column];

                    // Copy the value and formatting from the source cell to the destination cell
                    destinationCell.Value = sourceCell.Value;
                    sourceCell.Copy(destinationCell);
                }
            }

        }

        private void PrintData(Worksheet ws)
        {
            // Get the range of cells containing data
            Range usedRange = ws.UsedRange;

            // Loop through the cells and print the data
            for (int row = 1; row <= usedRange.Rows.Count; row++)
            {
                for (int column = 1; column <= usedRange.Columns.Count; column++)
                {
                    // Get the value of the current cell
                    Range cell = usedRange.Cells[row, column];
                    string cellValue = (cell.Value != null) ? cell.Value.ToString() : "";

                    // Print the value to the console
                    Console.Write(cellValue + "\t");
                }

                // Move to the next row
                Console.WriteLine();
            }
        }

        private async Task<bool> OpenCSVFiles(string csvFiles, Workbook wbHandle)
        {
            Worksheet worksheet = wbHandle.Worksheets.Add();

            try
            {
                // Import data from the CSV file
                string connectionString = "TEXT;" + csvFiles;
                QueryTable queryTable = worksheet.QueryTables.Add(connectionString, worksheet.Range["A1"]);
                queryTable.FieldNames = true;
                queryTable.RowNumbers = false;
                queryTable.FillAdjacentFormulas = false;
                queryTable.PreserveFormatting = true;
                queryTable.RefreshOnFileOpen = false;
                queryTable.RefreshStyle = XlCellInsertionMode.xlInsertDeleteCells;
                queryTable.SavePassword = false;
                queryTable.SaveData = true;
                queryTable.AdjustColumnWidth = true;
                queryTable.RefreshPeriod = 0;
                queryTable.TextFilePromptOnRefresh = false;
                queryTable.TextFilePlatform = (int)XlPlatform.xlWindows;
                queryTable.TextFileStartRow = 1;
                queryTable.TextFileParseType = XlTextParsingType.xlDelimited;
                queryTable.TextFileTextQualifier = XlTextQualifier.xlTextQualifierNone;
                queryTable.TextFileConsecutiveDelimiter = false;
                queryTable.TextFileTabDelimiter = true;
                queryTable.TextFileSemicolonDelimiter = false;
                queryTable.TextFileCommaDelimiter = false;
                queryTable.TextFileSpaceDelimiter = false;
                queryTable.TextFileColumnDataTypes = new object[] { Type.Missing };
                queryTable.TextFileTrailingMinusNumbers = true;
                queryTable.Refresh(false);
            }catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }

            return true;
        }

        private async Task<bool> Appending(Workbook wbHandle)
        {
            int workSheetCount = wbHandle.Worksheets.Count; 
            if(workSheetCount == 0)
            {
                return false;
            }
            int deltaPercentage = 50/ workSheetCount;
            List<int> indexList = new List<int>();
            for (int i=0;i<workSheetCount; i++)
            {
                indexList.Add(i+1);
            }
            List<int> mergedIndex = new List<int>();
            while(indexList.Count > 1)
            {
                for(int i = 0; i < indexList.Count-1; i += 2)
                {
                    if(! await Merged(wbHandle, indexList[i], indexList[i + 1]))
                    {
                        return false;
                    }
                    mergedIndex.Add(i+1);
                    if(i+3 >= indexList.Count)
                    {
                        foreach(int removedIdx in mergedIndex)
                        {
                            indexList.RemoveAt(removedIdx);
                        }
                        mergedIndex.Clear();
                        break;
                    }
                }
            }

            if(indexList.Count == 1) {
                return true;
            }
            return false;
        }

        private async Task<bool> Merged(Workbook wbHandle, int sheet1Index, int sheet2Index)
        {

            try
            {
                Worksheet worksheet1 = wbHandle.Worksheets[sheet1Index];
                Worksheet worksheet2 = wbHandle.Worksheets[sheet2Index];

                // Get the last used row in worksheet1
                int lastUsedRow = worksheet1.Cells.Find("*", System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious, false,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                // Get the range of data to copy from worksheet2
                Range rangeToCopy = worksheet2.UsedRange;

                // Paste the range into worksheet1
                rangeToCopy.Copy(worksheet1.Cells[lastUsedRow + 1, 1]);
            }catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
            return true;
        }

        private async Task<bool> SaveFile(Workbook wbResult,Workbook wbHandle, string filePath)
        {
            if(File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            Worksheet worksheet = wbResult.Worksheets.Add();

            wbHandle.Worksheets[1].Copy(worksheet);
            wbResult.Save();
            // Save the workbook
            wbResult.SaveAs(filePath);

            return true;
        }

        protected virtual void OnStatusChanged() {
            StatusChanged?.Invoke(this, EventArgs.Empty);
        }

        protected virtual void OnShowMessage() {
            ShowMessage?.Invoke(this, Message);
            mMessage = "";
        }
    }
}
