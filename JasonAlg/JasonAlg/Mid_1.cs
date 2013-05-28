using System;
using System.IO;
using System.Collections.Generic;

using System.Text;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;

namespace JasonAlg
{
    public class Mid_1
    {
        public int[] mid1;
        public int maxSPC;
        public int sumofSPC;
        private RawData rd;
        private StreamWriter sw;
        public Mid_1(ref RawData rd)
        {
            int[] spc;
            IComparer myComparer = new myReverserClass();
            spc= new int[rd.Cols];
            mid1 = new int[rd.Cols];
            this.rd = rd;
            sw = new StreamWriter(Common.debugpath+"\\"+"MID1.txt");
            Array.Copy(rd.spc, spc, rd.Cols);
            Array.Sort(spc, myComparer);
            Array.Clear(mid1, 0, mid1.Length);
            int val = spc[rd.tpsn-1+2];
            sumofSPC = 0;
            for (int i = 0; i < spc.Length; i++)
            {
                if (rd.spc[i] >= val) mid1[i] = 1;
                sumofSPC = sumofSPC + rd.spc[i];
            }
            maxSPC = spc[0];
            printSPC();

            printInputValues();
        }
        public void Close()
        {
            sw.Close();
        }
        private void printSPC()
        {
            sw.WriteLine("Mid_1(RawData.SPC) Print ------------------------------------");
            Common.printArray("MID1 :(SPC, sum of 1's ) ", rd.spc, ref sw);
            sw.WriteLine("Mid_1(Max SPC value) : " + maxSPC);
            sw.WriteLine("Mid_1(sum of SPC) : " + sumofSPC);
            sw.Write("Mid_1(Max PN's) : ");
            for (int i = 0; i < rd.spc.Length; i++)
            {
                if (rd.spc[i] >= maxSPC) sw.Write((i + 1).ToString() + " ");
            }
            sw.WriteLine();
            sw.WriteLine("Mid_1(RawData %SPC) Print ------------------------------------");
            Common.printArray("MID1 :(%SP) ", rd.PerSP, ref sw);
            sw.WriteLine("MID1 :(A%SP):{0,10}", rd.APerSP.ToString("######0.00"));
            sw.WriteLine();
            Common.printRankArray("MID1 : Rank ", rd.spc, ref sw);
            sw.WriteLine();
        }
        public void printInputValues()
        {

            // Print Raw Data

            sw.WriteLine();
            sw.WriteLine();
            sw.WriteLine();
            sw.WriteLine("Mid_1(RawData Print) ------------------------------------");
            //sw.Write("POSN");
            /*
            for (int i = 0; i < rd.Cols; i++)
                sw.Write("{0:4}",(i + 1).ToString());
            sw.WriteLine();
            */

            for (int i = 0; i < rd.Rows; i++)
            {
                if (i % 20 == 0)
                {
                    sw.Write("POSN");
                    for (int k = 0; k < rd.Cols; k++)
                        sw.Write("{0,4}", (k + 1).ToString());
                    sw.WriteLine();
                }
                sw.Write("{0,4}", (i + 1).ToString());
                for (int j = 0; j < rd.Cols; j++)
                {
                    if (rd.rawdata[i, j]==0)
                        sw.Write("____");
                    else
                        sw.Write("{0,4}", rd.rawdata[i, j].ToString());
                }
                sw.WriteLine();
                if (i % 100 == 0)
                    Console.WriteLine(" MID1: Print Line: " + i.ToString());
            }
            sw.WriteLine();
            
        }
        public void printMid1()
        {
            sw.WriteLine("Mid_1 Print ------------------------------------");
            Common.printArray("MID1 :(Mid_1) ",  mid1, ref sw);
            sw.WriteLine();
        }
        public void printToExcel(string path1, string sheetname)
        {
            ExcelWrapper ew = new ExcelWrapper();
            if (ew.Open(path1) == false) return;
            Excel.Worksheet mySheet;
            mySheet = ew.CreateSheet(sheetname);
            if (mySheet == null) { ew.Close(); return; }
            mySheet.Range["A1:HA200"].Clear();
            mySheet.Range["A1:AZ4000"].Clear();
            ew.printLine();

            ew.printLine("Mid_1(RawData.SPC) Print ------------------------------------");
            ew.printArray("MID1 :(SPC, sum of 1's ) ", rd.spc);
            ew.printLine("Mid_1(Max SPC value) : " + maxSPC);
            ew.printLine("Mid_1(sum of SPC) : " + sumofSPC);
            string pns = "";
            for (int i = 0; i < rd.spc.Length; i++)
            {
                if (rd.spc[i] >= maxSPC) pns = pns +(i + 1).ToString()+" ";
            } 
            ew.print("Mid_1(Max PN's) : "+pns);
            ew.printLine();
            ew.printLine("Mid_1(RawData %SPC) Print ------------------------------------");
            ew.printArray("MID1 :(%SP) ", rd.PerSP);
            ew.printLine("MID1 :(A%SP):"+ rd.APerSP.ToString("######0.00"));
            ew.printLine();
            ew.printRankArray("MID1 : Rank ", rd.spc);
            ew.printLine();

            ew.printArray("MID1 :(Mid_1) ", mid1);
            ew.printLine();

            // Print Raw Data
            /*
            ew.printLine();
            ew.printLine("Mid_1(RawData Print) ------------------------------------");
            ew.print("POSN");
            for (int i = 0; i < rd.Cols; i++) 
                ew.print((i + 1).ToString());
            ew.printLine();

            for (int i = 0; i < rd.Rows; i++)
            {
                ew.print((i + 1).ToString());
                for (int j = 0; j < rd.Cols; j++)
                {
                    if (rd.rawdata[i, j]==0)
                        ew.print(" ");
                    else
                        ew.print(rd.rawdata[i, j].ToString());
                }
                ew.printLine();
                if (i % 100 == 0)
                    Console.WriteLine(" MID1: Print Line: " + i.ToString());
            }
            ew.printLine();
            */
            ew.Save();
            ew.Close();
        }
    }
}
