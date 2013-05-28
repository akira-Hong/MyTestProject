using System;
using System.IO;
using System.Collections.Generic;
using System.Collections;

using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace JasonAlg
{
    public class ResultDB
    {
        private Form1 form;
        private int[,] rdb;
        private int Rows, Cols;
        private int[] rsum;
        private int[] rankSum;
        private int ThreshVal;
        private StreamWriter sw;
        public ResultDB(Form1 f)
        {
            form = f;
            sw = new StreamWriter(Common.debugpath + "\\" + "ResultDB.txt");
            Cols = f.rd.Cols;
            rdb = new int[20,Cols];
            for (int i = 0; i < 20; i++) for (int j = 0; j < f.rd.Cols; j++) rdb[i, j] = 0;
            Rows = 0;
            for (int j = 0; j < Cols; j++) rdb[Rows, j] = f.md1.mid1[j];
            Rows++;
            for (int j = 0; j < Cols; j++) rdb[Rows, j] = f.md2.mid2[j];
            Rows++;
            for (int k = 0; k < f.md3_1.mid3_1_Count; k++)
            {
                for (int j = 0; j < Cols; j++) rdb[Rows, j] = f.md3_1.mid3_1[k,j];
                Rows++;
            }
            for (int j = 0; j < Cols; j++) rdb[Rows, j] = f.md3_2.mid3_2[j];
            Rows++;
            for (int j = 0; j < Cols; j++) rdb[Rows, j] = f.md3_4.mid3_4[j];
            Rows++;
            for (int j = 0; j < Cols; j++) rdb[Rows, j] = f.md4_1.mid4_1[j];
            Rows++;
            for (int j = 0; j < Cols; j++) rdb[Rows, j] = (f.md4_3.mid4_3[j]>0)?1:0;
            Rows++;
            // sum
            int sum;
            rsum = new int[Cols];
            int [] temp = new int[Cols];
            for (int i = 0; i < Cols; i++)
            {
                sum = 0;
                for (int j = 0; j < Rows; j++)
                    if (rdb[j, i] > 0) sum++;
                rdb[Rows, i] = sum;
                rsum[i] = sum;
                temp[i] = sum;
            }
            Rows++;

            Common.printArray("MID-1 ", form.md1.mid1, ref sw);
            Common.printArray("MID-2 ", form.md2.mid2, ref sw);
            int[] temp2 = new int[Cols];
            for (int k = 0; k < f.md3_1.mid3_1_Count; k++)
            {
                for (int i = 0; i < Cols; i++) temp2[i] = form.md3_1.mid3_1[k, i];
                if (f.md3_1.mid3_1_Count==1)
                    Common.printArray("MID3-1 ", temp2, ref sw);
                else
                    Common.printArray("MID3-1-"+(k+1).ToString(), temp2, ref sw);
            }
            Common.printArray("MID3-2 ", form.md3_2.mid3_2, ref sw);
            Common.printArray("MID3-4 ", form.md3_4.mid3_4, ref sw);
            Common.printArray("MID4-1 ", form.md4_1.mid4_1, ref sw);
            Common.printArray("MID4-3 ", form.md4_3.mid4_3, ref sw);
            sw.WriteLine("");
            Common.printArray("ResultDB: SUM ", rsum, ref sw);

            // choose tspn + 2 PN's
            IComparer myComparer = new myReverserClass();
            Array.Sort(temp, myComparer);
            //Common.printArray("temp ", temp, ref sw);
            ThreshVal = temp[f.rd.tpsn - 1 + 2];

            sw.WriteLine("");
            sw.WriteLine((f.rd.tpsn + 2).ToString()+"th biggist value is " + ThreshVal.ToString());
            sw.WriteLine("");

            printTable(ref sw);

        }
        public void printTable(ref StreamWriter sw)
        {

            sw.WriteLine("RESULT DB TABLE Print ------------------------------------");
            for (int i = 0; i < Rows; i++)
            {
                for (int j = 0; j < Cols; j++)
                {
                    sw.Write("{0,4}", rdb[i, j].ToString("##0"));
                }
                sw.WriteLine();
            }
            sw.WriteLine();
            sw.WriteLine("RESULT DB : PN's List ------------------------------------");
            for (int i = 0; i < Cols; i++)
            {
                if (rsum[i] >= ThreshVal)
                    sw.Write("{0,4}", (i+1).ToString("##0"));
            }
            sw.WriteLine();
        }
        public void Close()
        {
            sw.Close();
        }
        public class myReverserClass : IComparer
        {
            // Calls CaseInsensitiveComparer.Compare with the parameters reversed.
            int IComparer.Compare(Object x, Object y)
            {
                return ((new CaseInsensitiveComparer()).Compare(y,x));
            }
        }
        public void printToExcel(string path1, string sheetname)
        {
            ExcelWrapper ew = new ExcelWrapper();
            if (ew.Open(path1) == false) return;
            Excel.Worksheet mySheet;
            mySheet = ew.CreateSheet(sheetname);
            if (mySheet == null) { ew.Close(); return; }
            mySheet.Range["A1:AZ40"].Clear();
            int[] temp = new int[Cols];
            for (int i = 0; i < Cols; i++) temp[i] = i + 1;
            ew.printArrayLine("PN", temp);
            ew.printArrayLine("MID-1 ", form.md1.mid1);
            ew.printArrayLine("MID-2 ", form.md2.mid2);
            for (int k = 0; k < form.md3_1.mid3_1_Count; k++)
            {
                for (int i = 0; i < Cols; i++) temp[i] = form.md3_1.mid3_1[k, i];
                if (form.md3_1.mid3_1_Count == 1)
                    ew.printArrayLine("MID3-1", temp);
                else
                    ew.printArrayLine("MID3-1-"+(k+1).ToString(), temp);
            }
            ew.printArrayLine("MID3-2 ", form.md3_2.mid3_2);
            ew.printArrayLine("MID3-4 ", form.md3_4.mid3_4);
            ew.printArrayLine("MID4-1 ", form.md4_1.mid4_1);
            ew.printArrayLine("MID4-3 ", form.md4_3.mid4_3);
            ew.printArrayLine("SUM ", rsum,1);
            ew.printLine("");
            string biggiest=(form.rd.tpsn + 2).ToString() + "th biggist value is " + ThreshVal.ToString();
            for (int i = 0; i < Cols; i++) temp[i] = 0;
            int k1 = 0;
            for (int i = 0; i < Cols; i++)
            {
                if (rsum[i] >= ThreshVal)
                    temp[k1++] = (i + 1);
            }
            ew.printArrayLine(biggiest, temp);
            ew.printLine();

            // print : pn list
            temp = new int[Cols];
            for (int i = 0; i < Cols; i++) temp[i] = i + 1;
            ew.printArrayLine("PN", temp);
            // print rank of totalsum
            rankSum = new int[rsum.Length];
            Common.makeRankArray(ref rsum, ref rankSum);
            ew.printArrayLine("Total_Sum", rankSum);

            // Rank Print
            // Mid 1
            //ew.print("MID 1");
            //ew.printRankArray("MID1:Rank", form.rd.spc,1);
            //ew.printLine();
            int[] mid1Rank = new int[Cols];
            Common.makeRankArray(ref form.rd.spc, ref mid1Rank);
            ew.printArrayLine("MID-1", mid1Rank);
            // MID 3-1
            ew.printArrayLine("MID3-1", form.md3_1.rank);
            int[] mid1_mid3_1 = new int[Cols];
            for (int i = 0; i < Cols; i++)
                mid1_mid3_1[i] = mid1Rank[i] + form.md3_1.rank[i];
            ew.printArrayLine("MID-1 & MID3-1", mid1_mid3_1);
            ew.printLine();


            // Mid 2
            //ew.print("MID 2");
            //ew.printRankArray("MID2:Rank (sort by PN's count in the List)", form.md2.temp, 1);
            //ew.printLine();
            int[] mid2Rank = new int[Cols];
            Common.makeRankArray(ref form.md2.temp, ref mid2Rank);
            ew.printArrayLine("MID-2", mid2Rank);

            // Mid3-4
            ew.printArrayLine("MID3-4", form.md3_4.rank3_4);
            int[] mid2_mid3_4 = new int[Cols];
            for (int i = 0; i < Cols; i++)
                mid2_mid3_4[i] = mid2Rank[i] + form.md3_4.rank3_4[i];
            ew.printArrayLine("MID-2 & MID3-4", mid2_mid3_4);
            ew.printLine();

            ew.printArrayLine("MID3-2 ", form.md3_2.mid3_2);
            ew.printArrayLine("MID4-1 ", form.md4_1.mid4_1);
            ew.printArrayLine("MID4-3 ", form.md4_3.mid4_3);
            int[] mid3_2mid4_1mid4_3 = new int[Cols];
            for (int i = 0; i < Cols; i++)
                mid3_2mid4_1mid4_3[i] = form.md3_2.mid3_2[i] + form.md4_1.mid4_1[i] + form.md4_3.mid4_3[i];
            ew.printArrayLine("MID3-2,Mid4-2 & Mid4-3", mid3_2mid4_1mid4_3,1);
            ew.printLine();
            //print final result
            int[] finalresult = new int[Cols];
            for (int i = 0; i < Cols; i++)
                finalresult[i] = rankSum[i] + mid1_mid3_1[i] + mid2_mid3_4[i] + mid3_2mid4_1mid4_3[i];
            ew.printArrayLine("FINAL RESULT", finalresult, 1);

            // print pn
            for (int i = 0; i < Cols; i++) temp[i] = i + 1;
            ew.printArrayLine("PN", temp);

            // Final Sort
            int [] finalsort=new int[Cols];
            int[] finalorder = new int[Cols];
            Array.Copy(finalresult, temp, Cols);
            Array.Sort(temp);
            Array.Clear(finalorder, 0, Cols);
            int prank=1, rank=1,cur=0;
            for (int i = 0; i < Cols; i++)
            {
                if (i > 0)
                {
                    if (temp[i] == temp[i - 1]) continue;
                }
                prank=rank;
                for (int j = 0; j < Cols; j++)
                {
                    if (finalresult[j] == temp[i]) { finalsort[j] = prank; finalorder[cur++] = j+1;   rank++; }
                }
            }
            ew.printArrayLine("Final Rank", finalsort,1);
            ew.printLine();
            ew.printArrayLine("Final PN's Order", finalorder,1);

            ew.Save();
            ew.Close();
        }
    }
}
