﻿using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Append_Excel
{
    class AppendHandler
    {
        public event EventHandler StatusChanged;
        public event EventHandler<string> ShowMessage;
        private static object mLocker = new object();

        private string mTemplatesPath = Environment.GetFolderPath(Environment.SpecialFolder.Templates);


        private bool mIsFailedBecauseOfTooLong = false;
        private bool mIsProcessing = false;
        private int mPercentageProcess = 0;
        private int mExecutedTime = 0;
        private int mEstimateTime = 0;
        private DateTime mStartTime = DateTime.MinValue;


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
                lock(mLocker)
                {
                    if (value != mPercentageProcess)
                    {
                        mPercentageProcess = value;
                        OnStatusChanged();
                        TimeEstimating();
                    }
                }
            } 
            get { return mPercentageProcess; } 
        }

        private string mMessage = "";
        public string Message
        {
            private set
            {
                if(value == "" || value != mMessage)
                {
                    mMessage = value;
                    OnShowMessage();
                }
            }
            get { return mMessage; }
        }

        public int ExecutedTime
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
                if(value < 1000)
                {
                    mEstimateTime = 1000;
                    OnStatusChanged();
                }else if(value != mEstimateTime)
                {
                    mEstimateTime = value; 
                    OnStatusChanged();
                }
            }
            get { return mEstimateTime; } 
        }
        public AppendHandler() { }


        public async Task StartProcessing(List<string> selectedFiles,string savePath, string sheetName)
        {
            if(selectedFiles.Count == 0)
            {
                Message = "Has no file selected";
                return;
            }
            Message = "";
            EstimatedTime = 0;
            ExecutedTime = 0;

            mIsFailedBecauseOfTooLong = false;
            IsProcessing = true;
            //handling
            mStartTime = DateTime.Now;
            try
            {
                //Console.WriteLine("tracking " + (DateTime.Now - start).TotalMilliseconds);
                mExcel = new Application();
                var ext = System.IO.Path.GetExtension(savePath);
                if (ext == ".csv")
                {
                    List<string> templateFiles = new List<string>();
                    int xlsxCount = 0, csvCount = 0;
                    foreach (string filePath in selectedFiles)
                    {
                        //Console.WriteLine("tracking " + (DateTime.Now - start).TotalMilliseconds);
                        var fileExt = Path.GetExtension(filePath);
                        if (fileExt == ".xlsx" || fileExt == ".xls")
                        {
                            xlsxCount++;
                        }
                        else if (fileExt == ".csv")
                        {
                            csvCount++;
                        }
                    }

                    int openXlsxDelta = (int)(50 * 0.75 / xlsxCount);
                    int csvDelta = (100-openXlsxDelta*xlsxCount)/selectedFiles.Count;

                    List<string> appendFile = new List<string>();
                    
                    for (int i=0;i<selectedFiles.Count;i++)
                    {
                        Message = "Reading xlsx files... ("+(i+1)+"/"+selectedFiles.Count+")";
                        string filePath = selectedFiles[i];
                        //Console.WriteLine("tracking " + (DateTime.Now - start).TotalMilliseconds);
                        var fileExt = Path.GetExtension(filePath); 
                        if (fileExt == ".xlsx" || fileExt == ".xls")
                        {
                            var result = await OpenXLSX(filePath, sheetName);
                            if(result != null && result.Count > 0)
                            {
                                templateFiles.AddRange(result);
                            }
                            PercentageProcess += openXlsxDelta;
                        }else
                        {
                            appendFile.Add(filePath);
                        }
                    }

                    appendFile.AddRange(templateFiles);

                    if (File.Exists(savePath))
                    {
                        File.Delete(savePath);
                    }


                    
                    for (int i=0;i<appendFile.Count;i++)
                    {
                        Message = "Appending csv file to " + savePath + " : " + (i + 1) + "/" + appendFile.Count;
                        if (!await AppendCsvToFile(appendFile[i], savePath, i==0)) {
                            Message = "Append Csv to "+savePath+" failed!";
                            IsProcessing = false;
                            mExcel.Quit();
                            return;
                        }
                        PercentageProcess += csvDelta;
                    }

                    foreach(string filePath in templateFiles)
                    {
                        //Console.WriteLine("tracking " + (DateTime.Now - start).TotalMilliseconds);
                        var fileExt = Path.GetExtension(filePath);
                        if (fileExt == ".csv")
                        {
                            if (File.Exists(filePath))
                            {
                                File.Delete(filePath);
                            }
                            PercentageProcess += csvDelta;
                        }
                    }

                    //if(totalData.Count > 0)
                    //{
                    //Console.WriteLine("tracking " + (DateTime.Now - start).TotalMilliseconds);
                    Message = "Saved file to " + savePath;
                    IsProcessing = false;
                    mExcel.Quit();
                    PercentageProcess = 100;
                    //Console.WriteLine("tracking " + (DateTime.Now - start).TotalMilliseconds);
                    TimeEstimating();
                    return;
                    //}
                }
                else
                {
                    //Console.WriteLine("tracking " + (DateTime.Now - start).TotalMilliseconds);
                    Workbook wbResult = mExcel.Workbooks.Add();
                    Workbook wbHandle = mExcel.Workbooks.Add();
                    //Console.WriteLine("tracking " + (DateTime.Now - start).TotalMilliseconds);
                    await OpenDataFiles(selectedFiles, wbHandle,sheetName);
                    //Console.WriteLine("tracking " + (DateTime.Now - start).TotalMilliseconds);
                    if (await Appending(wbHandle))
                    {
                        //Console.WriteLine("tracking " + (DateTime.Now - start).TotalMilliseconds);
                        await SaveFile(wbResult, wbHandle, savePath);
                        Message = "Saved file to " + savePath;
                        wbResult.Close(true);
                        wbHandle.Close(false);
                        mExcel.Quit();
                        IsProcessing = false;
                        TimeEstimating() ;
                        //Console.WriteLine("tracking " + (DateTime.Now - start).TotalMilliseconds);
                        return;
                    }
                    wbResult.Close(false);
                    wbHandle.Close(false);
                }
                mExcel.Quit();
            }
            catch(Exception ex)
            {
                Message=ex.Message;
                IsProcessing = false;
                return;
            }
            IsProcessing = false;
            Message = "Append Failed"+(mIsFailedBecauseOfTooLong?". Please try save file as csv!":"");
        }

        private async Task<List<string>> OpenXLSX(string filePath,string sheetName)
        {
            // Loop through the rows and columns and put the data into a List<List<object>>
            List<string> tempFiles = new List<string>();
            Workbook workbook = mExcel.Workbooks.Open(filePath);
            Console.WriteLine(filePath + " WorkSheet Count : " + workbook.Worksheets.Count);
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                try
                {
                    if(sheetName == "" || worksheet.Name == sheetName)
                    {
                        // Set the CSV file format
                        XlFileFormat csvFormat = XlFileFormat.xlCSV;

                        string templateFile = Path.Combine(mTemplatesPath, "temp_" + Path.GetFileNameWithoutExtension(filePath)+".csv");
                        if (File.Exists(templateFile))
                        {
                            File.Delete(templateFile);
                        }
                        // Export the worksheet to a CSV file
                        worksheet.SaveAs(templateFile, csvFormat);
                        tempFiles.Add(templateFile);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    workbook.Close(false);
                    return new List<string>();
                }
            }
            workbook.Close(false);
            return tempFiles;
        }

        private async Task<List<List<object>>> OpenCSV(string filePath)
        {
            List<List<object>> data = new List<List<object>>();

            // Create the CSV configuration object
            var csvConfiguration = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                Delimiter = ",",
                IgnoreBlankLines = true,
                TrimOptions = TrimOptions.Trim
            };

            // Read the header row of the first CSV file
            using (var reader = new StreamReader(filePath))
            using (var csvReader = new CsvReader(reader, csvConfiguration))
            {
                csvReader.Read();
                csvReader.ReadHeader();

                var columnNames = csvReader.HeaderRecord;
                List<object> header = new List<object>();
                foreach(object col in columnNames)
                {
                   header.Add(col);
                }
                if(header.Count > 0)
                {
                    data.Add(header);
                }
                while (csvReader.Read())
                {
                    List<object> rowData = new List<object>();
                    foreach (var columnName in columnNames)
                    {
                        var value = csvReader.GetField(columnName);
                        rowData.Add(value);
                    }
                    data.Add(rowData);  
                }
            }

            Console.WriteLine("Read csv count : " + data.Count+" filenmame : "+filePath);    
            return data;
        }

        private async Task<bool> SaveCsv(List<List<object>> data, List<object> header, string savePath)
        {
            using (var writer = new StreamWriter(savePath))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                // Write the header
                foreach (string columnName in header)
                {
                    csv.WriteField(columnName);
                }
                csv.NextRecord();

                // Write the data
                foreach (List<object> row in data)
                {
                    foreach (object value in row)
                    {
                        csv.WriteField(value);
                    }
                    csv.NextRecord();
                }
            }
            return true;
        }

        public void TimeEstimating()
        {
            if (IsProcessing && PercentageProcess >= 0 && PercentageProcess < 100)
            {

                TimeSpan elapsed = DateTime.Now - mStartTime;

                if(PercentageProcess != 0)
                {
                    EstimatedTime = (int)((elapsed.TotalMilliseconds * 100) / PercentageProcess);
                }
                ExecutedTime = (int)elapsed.TotalMilliseconds;
                
                Console.WriteLine("percentage " + PercentageProcess + " executed " + ExecutedTime + " estimate " + EstimatedTime);
            }else
            {
                EstimatedTime = ExecutedTime;
            }
        }

        private async Task<bool> OpenDataFiles(List<string> selectedFiles, Workbook wbHandle, string sheetName)
        {
            int deltaPercentage = 60/selectedFiles.Count;
            foreach (string file in selectedFiles)
            {
                var fileExt = System.IO.Path.GetExtension(file);
                if (fileExt == ".xlsx" || fileExt == ".xls")
                {
                    try
                    {
                        await OpenXLSXFiles(file, wbHandle,sheetName);
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

        private async Task<bool> OpenXLSXFiles(string xlsxFiles, Workbook wbHandle,string sheetName)
        {
            Workbook workbook = mExcel.Workbooks.Open(xlsxFiles);
            Console.WriteLine(xlsxFiles+" WorkSheet Count : "+workbook.Worksheets.Count);
            foreach(Worksheet worksheet in workbook.Worksheets)
            {
                try
                {
                    if(sheetName != "" && worksheet.Name != sheetName)
                    {
                        continue;
                    }
                    //PrintData(worksheet);
                    if (wbHandle.Worksheets[1].Name != "Result_Sheet")
                    {
                        Copy(worksheet, wbHandle.Worksheets[1]);
                        wbHandle.Worksheets[1].Name = "Result_Sheet";
                    }else
                    {
                        // Add a new sheet to the destination workbook
                        Worksheet destinationSheet = wbHandle.Worksheets.Add(After: wbHandle.Worksheets[1]);

                        // Set the destination sheet name to match the source sheet name
                        destinationSheet.Name = worksheet.Name + wbHandle.Worksheets.Count;

                        // Copy the entire source sheet to the destination sheet
                        Copy(worksheet, destinationSheet);
                    }
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

            sourceRange.Copy(destinationSheet.Cells[1, 1]);

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
            Workbook wb = mExcel.Workbooks.Add();
            Worksheet worksheet = wb.Worksheets[1];

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
                queryTable.TextFileCommaDelimiter = true;
                queryTable.TextFileSpaceDelimiter = false;
                queryTable.TextFileColumnDataTypes = new object[] { XlColumnDataType.xlTextFormat };
                queryTable.TextFileTrailingMinusNumbers = true;
                queryTable.Refresh(false);


                //PrintData(worksheet);
                if (wbHandle.Worksheets[1].Name != "Result_Sheet")
                {
                    Copy(worksheet, wbHandle.Worksheets[1]);
                    wbHandle.Worksheets[1].Name = "Result_Sheet";
                }
                else
                {
                    // Add a new sheet to the destination workbook
                    Worksheet destinationSheet = wbHandle.Worksheets.Add(After: wbHandle.Worksheets[1]);

                    // Set the destination sheet name to match the source sheet name
                    destinationSheet.Name = worksheet.Name + wbHandle.Worksheets.Count;

                    // Copy the entire source sheet to the destination sheet
                    Copy(worksheet, destinationSheet);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                wb.Close(false);
                return false;
            }
            wb.Close(false);

            return true;
        }

        private async Task<bool> Appending(Workbook wbHandle)
        {
            //foreach(Worksheet ws in wbHandle.Worksheets)
            //{
            //    PrintData(ws);
            //}

            int workSheetCount = wbHandle.Worksheets.Count; 
            if(workSheetCount == 0)
            {
                return false;
            }
            int perLoop = (int)(5 / Math.Log(workSheetCount,2));
            List<int> indexList = new List<int>();
            for (int i=0;i<workSheetCount; i++)
            {
                indexList.Add(i+1);
            }
            List<int> mergedIndex = new List<int>();
            while(indexList.Count > 1)
            {
                int perSubLoop = perLoop / (indexList.Count / 2);
                for(int i = 0; i < indexList.Count-1; i += 2)
                {
                    PercentageProcess += perSubLoop;
                    if (! await Merged(wbHandle, indexList[i], indexList[i + 1]))
                    {
                        return false;
                    }
                    mergedIndex.Add(indexList[i+1]); //save value to delete
                    if(i+3 >= indexList.Count)
                    {
                        foreach(int removedIdx in mergedIndex)
                        {
                            indexList.Remove(removedIdx);
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

        private async Task<bool> AppendCsvToFile(string sourceFile, string destinationFile, bool appendHeader = false)
        {
            // Create the CSV configuration object
            var csvConfiguration = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                Delimiter = ",",
                IgnoreBlankLines = true,
                TrimOptions = TrimOptions.Trim
            };

            // Read the header row of the first CSV file
            using (var reader = new StreamReader(sourceFile))
            // open the file for appending
            using (StreamWriter destinationWriter = new StreamWriter(destinationFile, true))
            using (var csvReader = new CsvReader(reader, csvConfiguration))
            {
                csvReader.Read();
                csvReader.ReadHeader();

                var columnNames = csvReader.HeaderRecord;
                if (appendHeader)
                {
                    string header = "";
                    for (int i = 0; i < columnNames.Length; i++)
                    {
                        header += columnNames[i];
                        if (i != columnNames.Length - 1) header += ",";
                    }
                    destinationWriter.WriteLine(header);
                }
                while (csvReader.Read())
                {
                    string newLine = "";
                    for (int i=0; i< columnNames.Length; i++)
                    {
                        var value = csvReader.GetField(columnNames[i]);
                        newLine += value;
                        if(i != columnNames.Length - 1)
                        {
                            newLine += ",";
                        }
                    }
                    if(newLine != "")
                    {
                        destinationWriter.WriteLine(newLine);
                    }
                }
            }

            return true;
        }

        private async Task<bool> Merged(Workbook wbHandle, int sheet1Index, int sheet2Index)
        {

            try
            {
                Worksheet worksheet1 = wbHandle.Worksheets[sheet1Index];
                Worksheet worksheet2 = wbHandle.Worksheets[sheet2Index];

                Console.WriteLine("Merged "+worksheet2.Name+" to "+worksheet2.Name + " with index "+sheet1Index+" and "+sheet2Index);
                Console.WriteLine("worksheet1 row : "+worksheet1.UsedRange.Rows.Count + " column "+worksheet1.UsedRange.Columns.Count);
                Console.WriteLine("worksheet2 row : " + worksheet2.UsedRange.Rows.Count + " column " + worksheet2.UsedRange.Columns.Count);

                // Get the last used row in worksheet1
                int lastUsedRow = worksheet1.Cells.Find("*", System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                    XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious, false,
                    System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                // Get the range of data to copy from worksheet2
                Range usedRange = worksheet2.UsedRange;
                Range rangeToCopy = usedRange.Offset[1, 0].Resize[usedRange.Rows.Count - 1, usedRange.Columns.Count];


                // Paste the range into worksheet1
                rangeToCopy.Copy(worksheet1.Cells[lastUsedRow + 1, 1]);

            }catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                if(ex.Message.Contains("aren't the same size"))
                {
                    mIsFailedBecauseOfTooLong = true;
                }
                return false;
            }
            return true;
        }

        private async Task<bool> SaveFile(Workbook wbResult,Workbook wbHandle, string filePath)
        {
            PercentageProcess = 65;
            if(File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            //foreach (Worksheet ws in wbHandle.Worksheets)
            //{
            //    PrintData(ws);
            //}
            PercentageProcess = 70;
            Copy(wbHandle.Worksheets[1], wbResult.Worksheets[1]);
            PercentageProcess = 80;
            //wbResult.Save();
            // Save the workbook
            wbResult.SaveAs(filePath);
            PercentageProcess = 100;

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
