using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using DataTable = System.Data.DataTable;
using System.IO;
using ExcelDataExtraction.Properties;

namespace ExcelDataExtraction
{
    class ExcelHelper
    {
        private string path;
        public Excel.Application excel = new Excel.Application();
        public Workbook wb;
        public Worksheet ws;
        public int sheetCount;
        public int columnCount;
        public int rowCount;
        public string worksheetname;



        //initialize the class
        public ExcelHelper(string path)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[1];
            sheetCount = wb.Worksheets.Count;
            columnCount = ws.UsedRange.Columns.Count;
            rowCount = ws.UsedRange.Rows.Count;
            worksheetname = ws.Name;
        }

        public void SaveWorkbook()
        {
            wb.Save();
        }

        public void CloseWorkbook()
        {
            wb.Close();
        }

        public void CloseApp()
        {
            excel.Quit();
        }
        public void OpenNewWorkbook(string path)
        {
            /*wb = excel.Workbooks.Open(path);*/
            wb = excel.Workbooks.Open(path, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
        }

        public void OpenWorkSheet(string sheet)
        {
            ws = wb.Worksheets[sheet];
            rowCount = ws.UsedRange.Rows.Count;
            worksheetname = ws.Name;
            columnCount = ws.UsedRange.Columns.Count;
        }

        public void WriteExcelCell(int row, int column, string data)
        {
            //string n = ws.Name;
            ((Range)ws.Cells[row, column]).Value = data;
        }
        public Dictionary<string, List<string>> ExtractExcelData(string filePath, string type)
        {
            //Start of Excel fcn
           
            Dictionary<string, List<string>> excelTempDict = new Dictionary<string, List<string>>();

            Excel.Workbook elWorkbook = wb;
            int MaxLimit = elWorkbook.Worksheets.Count;

            Excel.Worksheet workSheet = null;
            for (int ind = 1; ind <= MaxLimit; ind++)
            {
                workSheet = (Excel.Worksheet)elWorkbook.Sheets.Item[ind];
                if (type == "Read")
                {
                    workSheet.Select(Type.Missing);
                    excelTempDict = GetExcelData(workSheet);
                    // Write it to text file
                }
                else
                {
                    continue;
                }
            }
            elWorkbook.Close(0);
            excel.Quit();
            return excelTempDict;
        }
        public Dictionary<string, List<string>> GetExcelData(Excel.Worksheet workSheet1)
        {
            Dictionary<string, List<string>> ExcelDataDict = new Dictionary<string, List<string>>();

            Excel.Range last = workSheet1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range cellRange = workSheet1.Cells;
            cellRange.NumberFormat = "@";
            int lastUsedRow = last.Row;
            int lastUsedColumn = last.Column;

            for (int k = 1; k <= lastUsedColumn; k++)
            {
                string Onecolumn = k.ToString();
                List<string> rowsList = new List<string>();
                for (int i = 1; i <= lastUsedRow; i++)
                {
                    string CellValue = Convert.ToString((workSheet1.Cells[i, k] as Microsoft.Office.Interop.Excel.Range).Value);
                    rowsList.Add(CellValue);
                    
                }
                if (!ExcelDataDict.ContainsKey(Onecolumn))
                {
                    ExcelDataDict.Add(Onecolumn, rowsList);
                }
                else
                {
                    List<string> OldFinsdata = ExcelDataDict[Onecolumn];
                    OldFinsdata.AddRange(rowsList);
                    ExcelDataDict[Onecolumn] = OldFinsdata;
                }

            }
            return ExcelDataDict;
        }
     
        public void GenerateOutputText(Dictionary<string, List<string>> OutputData_Dict)
        {
            string OutputPath = Environment.CurrentDirectory;
            string timeStamp = string.Format("{0:dd_MM_yy_HH_mm_ss}", DateTime.Now).ToString();
            string OutputFilepath = Path.Combine(OutputPath, "ExcelToText_" + timeStamp + ".txt");
            string Concatenatedtext = "";
            foreach (var key in OutputData_Dict.Keys)
            {
                foreach (var valueItem in OutputData_Dict[key])
                {
                    Concatenatedtext= string.Join(",",Concatenatedtext, valueItem);
                   
                }
                Concatenatedtext = Concatenatedtext + Environment.NewLine;
            }
            File.WriteAllText(OutputFilepath, Concatenatedtext);

        }
    }
}
