using System;
using System.Collections.Generic;

using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;

namespace JasonAlg
{
    public class ExcelWrapper
    {
        private string filename;
        private Excel.Workbook excelWorkbook;
        private Excel.Application excelApp;
        private Excel.Sheets excelSheets;
        private Excel.Worksheet excelWorksheet;
        private bool IsOpen;
        private int curRow, curCol;
        public ExcelWrapper()
        {
            excelWorkbook = null;
            excelSheets = null;
            excelWorksheet = null;
            excelApp = new Excel.Application();
            IsOpen = false;
            curRow=curCol=1;
        }
        public bool Open(string filename)
        {
            this.filename = filename;
            excelApp.Visible = false;  // Makes Excel visible to the user.
            try
            {
                excelWorkbook = excelApp.Workbooks.Open(filename, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                IsOpen = true;
                excelSheets = excelWorkbook.Worksheets;
                excelWorksheet = null;
                curRow = curCol = 1;
                return true;
            }
            catch (Exception ex)
            {
                if (IsOpen == true) excelWorkbook.Close();
                Console.WriteLine("File open Error:  " + ex.Message);
            }
            IsOpen = false;
            return false;
        }
        public Excel.Worksheet FindSheet(string name)
        {
            if (!IsOpen) return null;
            foreach (Excel.Worksheet sh in excelSheets)
            {
                if (sh.Name == name) { excelWorksheet = sh; curRow = curCol = 1;  return sh; }
            }
            return null;
        }
        public Excel.Worksheet CreateSheet(string name)
        {
            if (!IsOpen) return null;
            Excel.Worksheet temp=FindSheet(name);
            if (temp!=null) return temp;
            try
            {
                temp = excelSheets.Add();
                temp.Name = name;
            }
            catch (Exception ex)
            {
                Console.WriteLine("ExcelWrapper CreateSheet: Error, " + ex.Message);
                temp = null;
            }
            excelWorksheet = temp;
            curRow = curCol = 1;
            return temp;
        }
        public void Save()
        {
            if (IsOpen)
            {
                excelWorkbook.Save();
            }
        }
        public void Close()
        {
            if (IsOpen)
            {
                excelWorkbook.Close();
            }
        }
        public void SetRowCol(int r, int c)
        {
            curRow = r;
            curCol = c;
        }
        public  void printArray(string header, int[] arr,int start=0)
        {
            try
            {
                curCol = 1 + start;
                excelWorksheet.Cells[curRow++, curCol] = "ARRAYPRINT: " + header;
                curCol = 1 + start;
                for (int i = 0; i < arr.Length; i++)
                {
                    excelWorksheet.Cells[curRow, curCol++] = (i + 1).ToString("####0");
                }
                curRow++;
                curCol = 1 + start;
                for (int i = 0; i < arr.Length; i++)
                {
                    excelWorksheet.Cells[curRow, curCol++] = arr[i].ToString("######0.00");
                }
                curRow++; curCol = 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("ExcelWrapper:printArray, ERROR: " + ex.Message);
            }
        }
        public  void printArray(string header, double[] arr,int start=0)
        {
            try{
                curCol = 1 + start;
                excelWorksheet.Cells[curRow++, curCol] = "ARRAYPRINT: " + header;
                curCol = 1 + start;
                for (int i = 0; i < arr.Length; i++)
                {
                    excelWorksheet.Cells[curRow,curCol++]=  (i + 1).ToString("####0");
                }
                curRow++; 
                curCol = 1 + start;
                for (int i = 0; i < arr.Length; i++)
                {
                    excelWorksheet.Cells[curRow,curCol++]=arr[i].ToString("####0.00");
                }
                curRow++; curCol = 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("ExcelWrapper:printArray2, ERROR: " + ex.Message);
            }
        }
        public void printArrayPercent(string header, int[] arr)
        {
            try
            {

                excelWorksheet.Cells[curRow++, curCol] = "ARRAYPRINT: " + header;
                for (int i = 0; i < arr.Length; i++)
                {
                    excelWorksheet.Cells[curRow, curCol++] ="'"+ ((i + 1)*Common.GpspInterval).ToString("####0.0");
                }
                curRow++; curCol = 1;
                for (int i = 0; i < arr.Length; i++)
                {
                    excelWorksheet.Cells[curRow, curCol++] = "'"+arr[i].ToString("####0.0");
                }
                curRow++; curCol = 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("ExcelWrapper:printArray, ERROR: " + ex.Message);
            }
        }
        public void printArrayLine(string header, int[] arr, int print0=0)
        {
            try
            {
                curCol = 1;
                excelWorksheet.Cells[curRow, curCol] = header;
                curCol++;
                for (int i = 0; i < arr.Length; i++)
                {
                    if (arr[i] == 0)
                    {
                        if (print0 == 0)
                            excelWorksheet.Cells[curRow, curCol++] = " ";
                        else
                            excelWorksheet.Cells[curRow, curCol++] = arr[i].ToString("####0");
                    }
                    else
                        excelWorksheet.Cells[curRow, curCol++] = arr[i].ToString("####0");
                }
                curRow++; curCol = 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("ExcelWrapper:printArrayLine, ERROR: " + ex.Message);
            }
        }
        public void printLine()
        {
            curRow++; curCol = 1;
        }
        public void printLine(string data)
        {
            try
            {
                excelWorksheet.Cells[curRow, curCol] = data;
                curRow++; curCol = 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("ExcelWrapper:printLine, ERROR: " + ex.Message);
            }
        }
        public void printListPercent(R1pnList rlist)
        {
            try
            {
                rlist.printPercentToExcel(this);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ExcelWrapper:printLine, ERROR: " + ex.Message);
            }
        }
        public void print()
        {
            curCol++;
        }
        public void print(string data)
        {
            try
            {
                excelWorksheet.Cells[curRow, curCol++] = data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("ExcelWrapper:print, ERROR: " + ex.Message);
            }
        }
        public  void printRankArray(string header, int[] arr,int kkk=0)
        {
            int[] spc = new int[arr.Length];
            int[] rank = new int[arr.Length];
            IComparer myComparer = new myReverserClass();
            Array.Copy(arr, spc, arr.Length);
            Array.Sort(spc, myComparer);
            int prev = Int32.MinValue;
            for (int i = 0; i < arr.Length; i++)
            {
                if (prev != spc[i])
                {
                    for (int j = 0; j < arr.Length; j++)
                    {
                        if (arr[j] == spc[i]) rank[j] = i + 1;
                    }
                    prev = spc[i];
                }
            }
            printArray(header, rank,kkk);
        }
    }
}
