using System;
using System.IO;
using System.Collections.Generic;

using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace JasonAlg
{
    public class Mid_3_2
    {
        private RawData rd;
        private Mid_3_2_Table t;
        private StreamWriter sw;
        public int[] mid3_2;
        private double intmin, intmax;
        public Mid_3_2(ref RawData rawdata)
        {
            rd = rawdata;
            t = new Mid_3_2_Table(rd.Rows, rd.tpsn);
            mid3_2 = new int[rd.Cols];
            Array.Clear(mid3_2, 0, mid3_2.Length);
            sw = new StreamWriter(Common.debugpath + "\\" + "MID3_2.txt");

            // make Mid3-2-table 
            int cindex = 0;
            for (int r = 0; r < rd.Rows; r++)
            {
                cindex=0;
                for (int c = 0; c < rd.Cols; c++)
                {
                    if (rd.rawdata[r,c]>0)
                    {
                        t.Set(r,cindex,rd.PerSP[c]);
                        cindex++;
                        if (cindex == rd.tpsn) break;
                    }
                }
            }

            // make AVERAGE
            t.makeAVERAGE();

            t.printTable(ref sw);

            // make mid3-2
            intmin = ((int)(t.maxavg / Common.GpspInterval)) * Common.GpspInterval;
            intmax = intmin + Common.GpspInterval;
            Array.Clear(mid3_2, 0, mid3_2.Length);
            for (int c = 0; c < rd.Cols; c++)
            {
                if (rd.PerSP[c] >= intmin && rd.PerSP[c] < intmax)
                {
                    mid3_2[c] = 1;
                }
            }
            sw.WriteLine("MID3_2_Table: interval lower : " + intmin.ToString());
            sw.WriteLine("MID3_2_Table: interval upper : " + intmax.ToString());

            Common.printArray("MID3_2 Array: ", mid3_2, ref sw);

        }
        public void Close()
        {
            sw.Close();
            t = null;
        }
        public void printToExcel(string path1, string sheetname)
        {
            ExcelWrapper ew = new ExcelWrapper();
            if (ew.Open(path1) == false) return;
            Excel.Worksheet mySheet;
            mySheet = ew.CreateSheet(sheetname);
            if (mySheet == null) { ew.Close(); return; }
            mySheet.Range["A1:HA200"].Clear();
/*
            ew.printLine("MID3_2_Table Print ------------------------------------");
            for (int i = 0; i < t.Rows; i++)
            {
                for (int j = 0; j < t.Cols; j++)
                {
                    ew.print(t.table[i, j].ToString("##0.0"));
                }
                ew.printLine();
            }
            ew.printLine();
 */ 
            ew.printLine("MID3_2_Table AVERAGE ------------------------------------");
            for (int j = 0; j < t.Cols; j++)
            {
                ew.print(t.avg[j].ToString("##0.00"));
            }
            ew.printLine();
            ew.printLine();
            // sw.WriteLine("MID3_2_Table: MaxIndex : "+maxindex.ToString());
            ew.printLine("MID3_2_Table: MaxAVG : " + t.maxavg.ToString("####0.00"));
            ew.printLine("MID3_2_Table: interval lower : " + intmin.ToString());
            ew.printLine("MID3_2_Table: interval upper : " + intmax.ToString());

            ew.printLine();
            ew.printArray("MID3_2: Result:", mid3_2);
            ew.printLine();
            ew.Save();
            ew.Close();
        }
    }
}
