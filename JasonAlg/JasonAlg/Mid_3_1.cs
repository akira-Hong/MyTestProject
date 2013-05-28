using System;
using System.IO;
using System.Collections.Generic;

using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace JasonAlg
{
    public class Mid_3_1
    {
        private const int max_mid31_Count = 10;
        private RawData rd;
        private R1pnList rlist;
        private StreamWriter sw;
        public int mid3_1_Count;
        public int[,] mid3_1;
        public int[] rank;
        public Mid_3_1(ref RawData r)
        {
            rd = r;
            rlist = new R1pnList();
            mid3_1 = new int[max_mid31_Count, rd.Cols];
            mid3_1_Count = 0;
            for (int i = 0; i < max_mid31_Count; i++) for (int j = 0; j < rd.Cols; j++) mid3_1[i, j] = 0;

            sw = new StreamWriter(Common.debugpath + "\\" + "MID3_1.txt");
            int count =(int)( 100 / Common.GpspInterval);
            for (int i = 0; i < count; i++)
            {
                R1pn r1 = new R1pn(5);
                rlist.Add(r1);
            }

            // make G%SP table
            int index;
            for (Int16 i = 0; i < r.Cols; i++)
            {
                index = (int)(r.PerSP[i] / Common.GpspInterval);
                rlist.rList[index].Add(i);
            }
            // test code
           // rlist.rList[3].Add(5);
           // rlist.rList[3].Add(9);
           // rlist.rList[3].Add(28);
           // rlist.rList[3].Add(45);
           // rlist.rList[3].Add(34);

            //Get Max index of a R1pn
            count = rlist.GetMaxCount();
            index = 0;
            for (int i = 0; i < rlist.rList.Count; i++)
            {
                if (rlist.rList[i].cp == count)
                {
                    for (int j = 0; j < rlist.rList[i].cp; j++)
                        mid3_1[index, rlist.rList[i].r1pn[j]] = 1;
                    index++;
                }
            }
            mid3_1_Count = index;

            rlist.print(ref sw, 10);

            int[] temp = new int[rd.Cols];
            if (mid3_1_Count == 1)
            {
                for (int i = 0; i < rd.Cols; i++) temp[i] = mid3_1[0, i];
                Common.printArray("Mid-3-1: Result: ", temp, ref sw);
            }
            else
            {
                for (int k = 0; k < mid3_1_Count; k++)
                {
                    for (int i = 0; i < rd.Cols; i++) temp[i] = mid3_1[k, i];
                    Common.printArray("Mid-3-1-"+(k+1).ToString()+": Result: ", temp, ref sw);
                }
            }

            rank = new int[rd.Cols];
            rlist.makeRankArray(ref rank);

        }
        public void Close()
        {
            sw.Close();
            rlist = null;
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
            int[] temp = new int[rd.Cols];
            if (mid3_1_Count == 1)
            {
                for (int i = 0; i < rd.Cols; i++) temp[i] = mid3_1[0, i];
                ew.printArray("Mid-3-1: Result: ", temp);
            }
            else
            {
                for (int k = 0; k < mid3_1_Count; k++)
                {
                    for (int i = 0; i < rd.Cols; i++) temp[i] = mid3_1[k, i];
                    ew.printArray("Mid-3-1-" + (k + 1).ToString() + ": Result: ", temp);
                }
            }
            ew.printLine();

            ew.printLine("MID3 RANK:");
            ew.printLine();
            ew.printArray("RANK:", rank);
            ew.printLine();

            ew.Save();
            ew.Close();
        }
    }
}
