using System;
using System.IO;
using System.Collections.Generic;

using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace JasonAlg
{
    public class Mid_3_3
    {
        public Mid_3_3_table mid3_3_table;
        public RawData rd;
        private StreamWriter sw;
        //public int[] mid3_3;
        public  Mid_3_3(ref RawData r)
        {
            rd = r;
            mid3_3_table = new Mid_3_3_table(ref r);
            //mid3_3 = new int[r.Cols];
            //Array.Clear(mid3_3, 0, mid3_3.Length);
            sw = new StreamWriter(Common.debugpath + "\\" + "MID3_3.txt");

            
            // get interval
            double intmin, intmax;
            intmax = (mid3_3_table.spcindex + 1) * Common.GpspInterval;
            intmin = intmax - Common.GpspInterval;

            // get PNs between intmax and intmin
            //for (int i = 0; i < r.Cols; i++)
           // {
           //     if (r.PerSP[i] > intmin && r.PerSP[i] <= intmax)
           ////         mid3_3[i] = 1;
           // }

            mid3_3_table.printtable(ref sw);
            sw.WriteLine("MID3_3_Table: interval lower : " + intmin.ToString("###,##0.0"));
            sw.WriteLine("MID3_3_Table: interval upper : " + intmax.ToString("###,##0.0"));

            //Common.printArray("MID3_3 Array: ", mid3_3, ref sw);
            
        }
        public void Close()
        {
            sw.Close();
        }
        public void printToExcel(string path1, string sheetname)
        {
            ExcelWrapper ew = new ExcelWrapper();
            if (ew.Open(path1) == false) return;
            Excel.Worksheet mySheet;
            mySheet = ew.CreateSheet(sheetname);
            if (mySheet == null) { ew.Close(); return; }
            mySheet.Range["A1:HA200"].Clear();
            ew.printLine();
            ew.printLine("MID3_3_Table SPC ------------------------------------");
            ew.print(" PN ");
            for (int j = 0; j < mid3_3_table.Cols; j++)
            {
                ew.print(((j + 1) * Common.GpspInterval).ToString("##0"));
            }
            ew.printLine();
            ew.print("SPC");
            for (int j = 0; j < mid3_3_table.Cols; j++)
            {
                ew.print(mid3_3_table.spc[j].ToString("##0"));
            }
            ew.printLine();
            ew.printLine();
            ew.printLine("MID3_3_Table spcIndex : " + mid3_3_table.spcindex.ToString() +", max PN : "+((mid3_3_table.spcindex+1)*Common.GpspInterval).ToString());
            ew.printLine("MID3_3_Table spcmax value   : " + mid3_3_table.spcmax.ToString());
            ew.printLine();
            ew.Save();
            ew.Close();
        }
    }
}
