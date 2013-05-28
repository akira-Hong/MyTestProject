using System;
using System.IO;
using System.Collections.Generic;

using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace JasonAlg
{
    public class Mid_4_2
    {
        private RawData rd;
        private StreamWriter sw;
        private Mid_4_1 md41;
        public int[,] Count_MxMn;
        public int[] Sumval;
        public double[] Avg;
        public int[] digit;
        public Mid_4_2(ref RawData raw, ref Mid_4_1 mid4_1)
        {
            rd = raw;
            this.md41 = mid4_1;
            sw = new StreamWriter(Common.debugpath + "\\" + "MID4_2.txt");

            //Make Count MX/MN Table 
            Count_MxMn = new int[rd.Rows, 4];
            Sumval = new int[4];
            Avg = new double[4];
            Array.Clear(Avg, 0, Avg.Length);
            digit = new int[4];
            Array.Clear(digit, 0, digit.Length);
            for (int i = 0; i < rd.Rows; i++) for (int j = 0; j < 4; j++) Count_MxMn[i, j] = 0;
            int Cols = 0,pn;
            foreach (R1pn r1 in md41.rlist.rList)
            {
                for (int i = 0; i < r1.cp; i++)
                {
                    pn=r1.r1pn[i];
                    for (int j = 0; j < rd.Rows; j++)
                        if (rd.rawdata[j, pn] > 0) Count_MxMn[j, Cols]++;
                }
                Cols++;
            }

            // Make Average
            double sum;
            for (int i = 0; i < 4; i++)
            {
                sum = 0;
                for (int j = 0; j < rd.Rows; j++)
                {
                    sum += Count_MxMn[j, i];
                }
                Sumval[i] =(int) sum;
                Avg[i] = sum / rd.Rows;
            }

            // get Digits 
            for (int i = 0; i < 4; i++)
                digit[i] =(int)( Avg[i] + 0.5);
            printTable(ref sw);

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
            string s = "A1:HA" + (Common.POSN + 100).ToString();
            mySheet.Range[s].Clear();
            ew.printLine();
            ew.printLine("MID4_2_Table (Count MX/MN Table) Print ------------------------------------");
            string [] h = new string [] {"POSN","MX+A%SP","MX-A%SP","MN+A%SP","MN-A%SP"};
            for (int i = 0; i <h.Length; i++)
            {
                ew.print(h[i]);
            }
            ew.printLine();
            for (int i = 0; i < rd.Rows; i++)
            {
                ew.print((i + 1).ToString("####0"));
                for (int j = 0; j < 4; j++)
                {
                    ew.print(Count_MxMn[i, j].ToString("####0"));
                }
                ew.printLine();
            }
            ew.printLine();
            //ew.printLine("MID4_2_Table sum ------------------------------------");
            ew.print("SUM");
            for (int j = 0; j < 4; j++)
            {
                ew.print(Sumval[j].ToString("####0"));
            }
            ew.printLine();
            //ew.printLine("MID4_2_Table AVERAGE ------------------------------------");
            ew.print("Average");
            for (int j = 0; j < 4; j++)
            {
                ew.print(Avg[j].ToString("##0.00"));
            }
            ew.printLine();
            //ew.printLine("MID4_2_Table DIGITs ------------------------------------");
            ew.print("DIGITs");
            for (int j = 0; j < 4; j++)
            {
                ew.print(digit[j].ToString("##0"));
            }
            ew.printLine();
            ew.printLine();
            ew.Save();
            ew.Close();
        }
        public void printTable(ref StreamWriter sw)
        {
            sw.WriteLine("MID4_2_Table (Count MX/MN Table) Print ------------------------------------");
            sw.WriteLine("POSN MX+A%SP MX-A%SP MN+A%SP MN-A%SP");
            for (int i = 0; i < rd.Rows; i++)
            {
                sw.Write("{0,4}", (i+1).ToString("####0"));
                for (int j = 0; j < 4; j++)
                {
                    sw.Write("{0,8}", Count_MxMn[i, j].ToString("####0"));
                }
                sw.WriteLine();
            }
            sw.WriteLine();
            sw.WriteLine("MID4_2_Table sum ------------------------------------");
            for (int j = 0; j < 4; j++)
            {
                sw.Write("{0,7}", Sumval[j].ToString("####0"));
            }
            sw.WriteLine();
            sw.WriteLine("MID4_2_Table AVERAGE ------------------------------------");
            for (int j = 0; j < 4; j++)
            {
                sw.Write("{0,7}", Avg[j].ToString("##0.00"));
            }
            sw.WriteLine();
            sw.WriteLine("MID4_2_Table DIGITs ------------------------------------");
            for (int j = 0; j < 4; j++)
            {
                sw.Write("{0,7}", digit[j].ToString("##0"));
            }
            sw.WriteLine();
        }
    }
}
