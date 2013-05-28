using System;
using System.IO;
using System.Collections.Generic;

using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace JasonAlg
{
    public class Mid_2
    {
        private RawData rd;
        private Mid_1 mid1;
        private int Step;
        private int MaxStep;
        private R1pnList rlist, rnext;
        private StreamWriter sw;
        public int[] temp;
        public int[] mid2;
        public Mid_2(ref RawData r, ref Mid_1 m1)
        {
            rd = r; mid1 = m1;
            mid2 = new int[rd.Cols];
            MaxStep = r.tpsn;
            rlist = new R1pnList();
            FirstMakeR1PN();
            sw = new StreamWriter(Common.debugpath + "\\" + "MID2.txt");
            printRList();
        }
        public void Close()
        {
            sw.Close();
            rlist = rnext = null;
        }
        public void printRList()
        {
            sw.WriteLine("MID2:printRList++++++++++++++++++++++++ ");
            foreach (R1pn r1 in rlist.rList)
            {
                r1.print(ref sw);
            }
            sw.WriteLine("+++++++++++++");
        }
        //
        private void FirstMakeR1PN()
        {
            for (Int16 i = 0; i < rd.Cols; i++)
            {
                if (rd.spc[i] == mid1.maxSPC)
                {
                    R1pn r1 = new R1pn(rd.tpsn);
                    r1.Add(i);
                    rlist.Add(r1);
                }
            }
        }

        // Make subset list from the R1pn.
        // for example,  (11,25, 39, 20)  => {(11,25,39,20), (25,39,20), (39,20), (20)}
        private void MakeSubsetofR1pn(ref R1pnList rsubset,int indexofbaseR1pn, R1pn r1)
        {
            int count = r1.cp;
            R1pn r2 = new R1pn(r1.r1pn.Length);
            r2.CopyFrom(r1);
            for (int i = 0; i < count; i++)
            {
               // if (!rsubset.IsExist(ref r2))  //check if rsubset already has the r2.
                {
                    R1pn new1 = new R1pn(r1.r1pn.Length);
                    new1.CopyFrom(r2);
                    new1.indexofbaseR1pn = indexofbaseR1pn;
                    rsubset.Add(new1);
                }
                r2.ShiftLeft_1();
                if (r2.cp <= 0) break;
            }
        }

        // Check if All values of r1 are '1', then return true
        private bool IsAllofR1pn_1(int index, R1pn r1)
        {
            for (int i = 0; i < r1.cp; i++)
            {
                if (rd.rawdata[index, r1.r1pn[i]] != 1) return false;
            }
            return true;
        }

        // process a subset.
        private void Find_Next_R1pn_Set(R1pn r1)
        {
            // make subtable and SPC
            int[] spc = new int[rd.Cols];
            Array.Clear(spc, 0, spc.Length);
            for (int i = 0; i < rd.Rows; i++)
            {
                if (IsAllofR1pn_1(i,r1))
                {
                    for (int j = 0; j < rd.Cols; j++)
                        if (rd.rawdata[i, j] == 1) spc[j]++;
                }
            }

            //Common.printArray("Find_Next_R1pn_Set:SPC(1)", ref spc, ref sw);

            // clear prev reference
            R1pn prev = rlist.rList[r1.indexofbaseR1pn];
            for (int i = 0; i < prev.cp; i++)
            {
                spc[prev.r1pn[i]] = 0;
            }

            //Common.printArray("Find_Next_R1pn_Set:SPC(2)", ref spc, ref sw);

            // get the spcs with the maximum value.
            int max = Common.MaxValue(ref spc);

            // make next R1pn with checking the same R1pn in the rlist.
            // put new R1pns' into the rnext list
            for (Int16 i = 0; i < spc.Length; i++)
            {
                if (spc[i] == max && max>=1)  // 중요한 부분임...   threshhold값에 따라 결과가 많이 달라짐.
                {
                    R1pn new1 = new R1pn(rd.tpsn);
                    new1.CopyFrom(prev);
                    new1.Add(i);
                    //if (!rnext.IsExist(ref new1))
                        rnext.Add(new1);
                }
            }
        }
        public void Run()
        {
            Array.Clear(mid2, 0, mid2.Length);
            for (Step = 0; Step < MaxStep; Step++)
            {
                Console.WriteLine("maxstep: {0}, current step:{1}", MaxStep, Step);
                rnext = null;
                rnext = new R1pnList();

                R1pnList rsubset = new R1pnList();
                sw.WriteLine("");
                sw.WriteLine("NEW R1PN Start: step:" + Step + " 's  R1pn List ");
                // make subset of the R1PN_X step.
                int indexofbaseR1pn = 0;
                foreach (R1pn r1 in rlist.rList)
                {
                    r1.print(ref sw);
                    MakeSubsetofR1pn(ref rsubset,indexofbaseR1pn, r1);
                    indexofbaseR1pn++;
                }

                GC.Collect();
                //sw.WriteLine("");
                //sw.WriteLine("rsubset's List  ");
                //rsubset.print(ref sw);

                // process each subset.
                int ccc = 0;
                foreach (R1pn r1 in rsubset.rList)
                {
                    Find_Next_R1pn_Set(r1);
                    ccc++;
                    if (ccc < 0) ccc = 0;
                    if (ccc % 100000 == 0)
                    {
                        Console.WriteLine("ccc: {0}", ccc);
                        GC.Collect();
                    }
                }

                //sw.WriteLine("");
                //sw.WriteLine("rnext's List  ");
                //rnext.print(ref sw);

                // Remove Current R1pn_X List
                // Set Next R1pn_X1 List
                rlist.rList = null;
                rlist.rList = rnext.rList;

                //Common.printArray("MID2 :("+Step.ToString()+") ",  mid2, ref sw);

                rnext = null;
                rsubset = null;
                GC.Collect();
            }

            // count each pn how many times it comes in all of rlist
            temp = new int[mid2.Length];
            rlist.FinalProcessing_PNS(ref temp);
            Common.printArray2("Temp :(temp,PN's count that is appeared in the List.) ",  temp, ref sw);
            // then, sort the count by pn
            // choose highest pn count as much as tspn number.
            int[] temp2 = new int[mid2.Length];
            Array.Copy(temp, temp2, temp.Length);
            Array.Sort(temp2);
           // Common.printArray2("Temp2 :(temp2, sort) ",  temp2, ref sw);
            Array.Clear(mid2, 0, mid2.Length);
            int val = 0, flag=0,count=0;
            for (int i = temp2.Length-1; i >=0; i--)
            {
                if (flag == 1  || temp2[i]==0) break;
                if (temp2[i] != val) 
                    val = temp2[i];
                else continue;
                for (int j = 0; j < temp.Length; j++)
                {
                    if (temp[j] == val)
                    {
                        mid2[j] = 1;
                        count++;
                    }
                    if (count == rd.tpsn + 2) { flag = 1; break; }
                }
            }

            sw.WriteLine("Choose  " + (rd.tpsn + 2).ToString("######0") +" PN's");
            sw.WriteLine("Threshold Count: " + (val).ToString("######0"));

            Common.printArray("MID2 :(LAST) ",  mid2, ref sw);
            Common.printIndex("MID2 INDEX Print :(LAST) ",  mid2, ref sw);
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
            ew.printArray("MID2 :(PN's Count) ", temp);
            ew.printLine();
            ew.printLine();
            ew.printArray("MID2 :(Mid_2) ", mid2);
            ew.printLine();
            ew.Save();
            ew.Close();
        }
    }
}
