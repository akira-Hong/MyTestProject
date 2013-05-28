using System;
using System.IO;
using System.Collections.Generic;

using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace JasonAlg
{
    public class Mid_4_3
    {
        private RawData rd;
        private Mid_4_2 md42;
        private StreamWriter sw;
        public int[] mid4_3;
        public Mid_4_3(ref RawData r, ref Mid_4_1 md41, ref Mid_4_2 m42)
        {
            rd = r;
            md42 = m42;
            mid4_3 = new int[rd.Cols];
            Array.Clear(mid4_3, 0, mid4_3.Length);
            sw = new StreamWriter(Common.debugpath + "\\" + "MID4_3.txt");

            // test ++++++++++++++++++++++++++++++++++++++++++++++++
           // md42.digit[0] = 1;
           // md42.digit[1] = 1;

            // make SPN Combination FROM Mid_4_2
            // make that each array in md41.rlist has multiple elements of digit[i].
            R1pnList rFrom, rTo;
            rFrom = md41.rlist.Clone();
            for (int i = 0; i < md42.digit.Length; i++)
            {
                if (md42.digit[i] <=1) continue;
                rFrom.rList[i].MakeMultiple(md42.digit[i]);
            }

            // make SPN Combination
            R1pnList rTarget = new R1pnList();
            rTo = new R1pnList();
            for (int i = 0; i < md42.digit.Length; i++)
            //for (int i = md42.digit.Length-1; i>=0 ; i--)
            {
                if (md42.digit[i] <= 0) continue;
                rTarget = new R1pnList();
                for (int j = 0; j < rFrom.rList[i].cp / md42.digit[i]; j++)
                {
                    R1pnList rToTo = rTo.Clone();
                    rToTo.AddAllList(rFrom.rList[i].r1pn, j * md42.digit[i], md42.digit[i]);
                    rTarget.MergeList(rToTo);
                    rToTo = null;
                }
                rTo = rTarget;
                GC.Collect();
            }
            rTo = null;
            rTarget.print(ref sw);

            // count each pn how many times it comes in the SPN Combination table.
            rTarget.getCountEachPN(ref mid4_3);
            Common.printArray("MID4_3 :MID4_3", mid4_3, ref sw);


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
            ew.printArray("MID4_3: Result:", mid4_3);
            ew.printLine();
            ew.Save();
            ew.Close();
        }
    }
}
