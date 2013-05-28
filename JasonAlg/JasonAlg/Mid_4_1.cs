using System;
using System.IO;
using System.Collections.Generic;

using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace JasonAlg
{
    // MX/MN_Deviation 
    public class Mid_4_1
    {
        private RawData rd;
        private StreamWriter sw;
        public double[] mx_mn;
        public R1pnList rlist;
        public int[] mid4_1;
        public Mid_4_1(ref RawData r)
        {
            rd = r;
            mid4_1 = new int[rd.Cols];
            mx_mn = new double[rd.Cols];
            rlist = new R1pnList();
            Array.Clear(mid4_1, 0, mid4_1.Length);
            sw = new StreamWriter(Common.debugpath + "\\" + "MID4_1.txt");

            Common.printArray("MID4_1 :Rawdata %SP ", rd.PerSP, ref sw);
            sw.WriteLine("MID4_1: Average %SP : " + rd.APerSP.ToString("#####0.00"));

            // make mx_mn array
            for (int i = 0; i < rd.Cols; i++)
            {
                mx_mn[i] = rd.APerSP - rd.PerSP[i];
            }
            // get the most number out of + values.
            int index = Common.MaxIndex(ref mx_mn);
            double val = mx_mn[index];
            R1pn r1 = new R1pn(10);
            for (Int16 i = 0; i < rd.Cols; i++) if (mx_mn[i] == val) { mid4_1[i] = 1; r1.Add(i); }
            rlist.Add(r1);
            Common.printArray("MID4_1 :first", mid4_1, ref sw);
            // get the least number out of - values;
            index = Common.MinIndex(ref mx_mn);
            val = mx_mn[index];
            r1 = new R1pn(10);
            for (Int16 i = 0; i < rd.Cols; i++) if (mx_mn[i] == val) { mid4_1[i] = 1; r1.Add(i); }
            rlist.Add(r1);
            Common.printArray("MID4_1 :second", mid4_1, ref sw);
            // get the least number out of + values;
            index = Common.MinIndexOutofPlus(ref mx_mn);
            val = mx_mn[index];
            r1 = new R1pn(10);
            for (Int16 i = 0; i < rd.Cols; i++) if (mx_mn[i] == val) { mid4_1[i] = 1; r1.Add(i); }
            rlist.Add(r1);
            Common.printArray("MID4_1 :third", mid4_1, ref sw);

            // get the most number out of - values;
            index = Common.MaxIndexOutofMinus(ref mx_mn);
            val = mx_mn[index];
            r1 = new R1pn(10);
            for (Int16 i = 0; i < rd.Cols; i++) if (mx_mn[i] == val) { mid4_1[i] = 1; r1.Add(i); }
            rlist.Add(r1);
            Common.printArray("MID4_1 :MX/MN", mx_mn, ref sw);
            Common.printArray("MID4_1 :MID4_1", mid4_1, ref sw);

            rlist.print(ref sw);

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
            ew.printArray("MID4_1 :Rawdata %SP ", rd.PerSP);
            ew.printLine("MID4_1: Average %SP : " + rd.APerSP.ToString("#####0.00"));
            ew.printLine();
            ew.printArray("MID4_1 :MX/MN", mx_mn);
            ew.printLine();

            // pn list 
            string[] header = new string[] { "MX+A%SP PN's List", "MX-A%SP PN's List", "MN+A%SP PN's List", "MN-A%SP PN's List" };
            int index = 0;
            foreach (R1pn r1 in rlist.rList)
            {
                ew.printLine(header[index]+" : "+r1.getString());
                index++;
                if (index == 4) break;
            }
            ew.printLine();

            ew.printArray("MID4_1: Result:", mid4_1);
            ew.printLine();

            ew.Save();
            ew.Close();
        }
    }
}
