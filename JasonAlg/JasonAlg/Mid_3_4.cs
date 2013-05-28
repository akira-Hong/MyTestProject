using System;
using System.IO;
using System.Collections.Generic;

using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace JasonAlg
{
    class Cpnlist
    {
        public int count, min, max;
        public int[] pns;
        public Cpnlist(int c)
        {
            count = 0;
            pns = new int[c];
        }
        public void Add(int val)
        {
            pns[count] = val;
            count++;
        }
    }
    public class Mid_3_4
    {
        private Mid_3_3_table m33;
        private RawData rd;
        private StreamWriter sw;
        private int[] temp, perTemp;
        public int[] mid3_4;
        public int[] rank3_4;
        private int maxSpcCount;

        private int Step;
        private int MaxStep;
        private R1pnList rlist, rnext;

        private int cplistcount;
        private Cpnlist[] cpnlist;

        public  Mid_3_4(ref Mid_3_3 m)
        {
            m33 = m.mid3_3_table;
            rd = m.rd;
            mid3_4 = new int[rd.Cols];
            Array.Clear(mid3_4, 0, mid3_4.Length);
            sw = new StreamWriter(Common.debugpath + "\\" + "MID3_4.txt");
            cplistcount = 0;
            cpnlist = new Cpnlist[m33.Cols];
            MaxStep = rd.tpsn;
            rlist = new R1pnList();
            FirstMakeR1PN();
            Run();

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
                r1.printPercent(ref sw);
            }
            sw.WriteLine("+++++++++++++");
        }
        //
        private void FirstMakeR1PN()
        {
            for (Int16 i = 0; i < m33.Cols; i++)
            {
                if (m33.spc[i] == m33.spcmax)
                {
                    R1pn r1 = new R1pn(rd.tpsn);
                    r1.Add(i);
                    rlist.Add(r1);
                }
            }
        }

        // Make subset list from the R1pn.
        // for example,  (11,25, 39, 20)  => {(11,25,39,20), (25,39,20), (39,20), (20)}
        private void MakeSubsetofR1pn(ref R1pnList rsubset, int indexofbaseR1pn, R1pn r1)
        {
            int count = r1.cp;
            R1pn r2 = new R1pn(r1.r1pn.Length);
            r2.CopyFrom(r1);
            for (int i = 0; i < count; i++)
            {
                if (!rsubset.IsExist(ref r2))  //check if rsubset already has the r2.
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
                if (m33.rawdata[index, r1.r1pn[i]] <=0) return false;
            }
            return true;
        }

        // process a subset.
        private void Find_Next_R1pn_Set(R1pn r1)
        {
            // make subtable and SPC
            int[] spc = new int[m33.Cols];
            Array.Clear(spc, 0, spc.Length);
            for (int i = 0; i < m33.Rows; i++)
            {
                if (IsAllofR1pn_1(i, r1))
                {
                    for (int j = 0; j < m33.Cols; j++)
                        if (m33.rawdata[i, j] > 0) spc[j] += m33.rawdata[i, j];
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
                if (spc[i] == max && max >= 1)  // 중요한 부분임...   threshhold값에 따라 결과가 많이 달라짐.
                {
                    R1pn new1 = new R1pn(rd.tpsn);
                    new1.CopyFrom(prev);
                    new1.Add(i);
                    if (!rnext.IsExist(ref new1))
                        rnext.Add(new1);
                }
            }
        }
        public void Run()
        {
            Array.Clear(mid3_4, 0, mid3_4.Length);
            for (Step = 0; Step < MaxStep; Step++)
            {
                rnext = null;
                rnext = new R1pnList();

                R1pnList rsubset = new R1pnList();
                sw.WriteLine("");
                sw.WriteLine("NEW R1PN Start: step:" + Step + " 's  R1pn List ");
                // make subset of the R1PN_X step.
                int indexofbaseR1pn = 0;
                foreach (R1pn r1 in rlist.rList)
                {
                    r1.printPercent(ref sw);
                    MakeSubsetofR1pn(ref rsubset, indexofbaseR1pn, r1);
                    indexofbaseR1pn++;
                }

                sw.WriteLine("");
                sw.WriteLine("rsubset's List  ");
                rsubset.printPercent(ref sw);

                // process each subset.
                foreach (R1pn r1 in rsubset.rList)
                {
                    Find_Next_R1pn_Set(r1);
                }

                sw.WriteLine("");
                sw.WriteLine("rnext's List  ");
                rnext.printPercent(ref sw);

                if (rnext.Count == 0)
                {
                    sw.WriteLine("rnext's List is empty ");
                    sw.WriteLine("exit this stage and go to final processing ");
                    rnext = null;
                    rsubset = null;
                    GC.Collect();
                    break;
                }

                // Remove Current R1pn_X List
                // Set Next R1pn_X1 List
                rlist.rList = null;
                rlist.rList = rnext.rList;

                //Common.printArray("MID3_4 :(" + Step.ToString() + ") ", mid2, ref sw);

                rnext = null;
                rsubset = null;
                GC.Collect();
            }

            // count each pn how many times it comes in all of rlist
            temp = new int[m33.Cols];
            rlist.FinalProcessing_PNS(ref temp);
            // then, sort the count by pn
            // choose highest pn count as much as tspn number.
            int[] temp2 = new int[m33.Cols];
            Array.Copy(temp, temp2, temp.Length);
            Array.Sort(temp2);
            Common.printArrayPercent("MID3-4 : %SP values (this means how many times does each %SP appear in the List)", temp, ref sw);
            //Common.printArrayPercent("Temp2 :(temp2) ", temp2, ref sw);

#if false
            int[] perTemp = new int[m33.Cols];
            Array.Clear(perTemp, 0, perTemp.Length);
            int val = 0, flag = 0, count = 0;
            for (int i = temp2.Length - 1; i >= 0; i--)
            {
                if (flag == 1 || temp2[i] == 0) break;
                if (temp2[i] != val)
                    val = temp2[i];
                else continue;
                for (int j = 0; j < temp.Length; j++)
                {
                    if (temp[j] == val)
                    {
                        perTemp[j] = 1;
                        count++;
                    }
                    if (count ==  rd.tpsn + 2 ) { flag = 1; break; }
                }
            }
#else
            perTemp = new int[m33.Cols];
            Array.Clear(perTemp, 0, perTemp.Length);
            for (int i = 0; i < temp.Length;i++ )
            {
                if (temp[i] == temp2[temp2.Length - 1]) perTemp[i] = 1;
            }
            maxSpcCount = temp2[temp2.Length - 1];

#endif
            sw.WriteLine();
            sw.WriteLine("max %SP value : " + maxSpcCount.ToString());
            sw.WriteLine();
            Common.printArrayPercent("MID3-4 :(Choose %SPs' that have max %SP value) ", perTemp, ref sw);

            // make mid3_4
            sw.WriteLine();
            //Common.printArray("RawData %SP: ", rd.PerSP, ref sw);
            //sw.WriteLine();

            cplistcount = 0;
            Array.Clear(mid3_4, 0, mid3_4.Length);
            for (int i = 0; i < perTemp.Length; i++)
            {
                if (perTemp[i] > 0)
                {
                    // get interval
                    double intmin, intmax;
                    intmax = (i + 1) * Common.GpspInterval;
                    intmin = intmax - Common.GpspInterval;
                    string pns = "";
                    pns = "Interval(%SP): " + (intmin).ToString("###0.0") + "% ~ " + intmax.ToString("###0.0") + "%  , ";
                    pns +=" PN's : ";
                    cpnlist[cplistcount] = new Cpnlist(rd.Cols);
                    // get PNs between intmax and intmin
                    for (int j = 0; j < rd.Cols; j++)
                    {
                        double  persp=rd.PerSP[j];
                        if (persp > intmin && persp <= intmax)
                        {
                            //mid3_4[j] = 1;
                            cpnlist[cplistcount].Add(j);
                            pns+=((j+1).ToString() + " ");
                        }
                    }
                    cplistcount++;
                    sw.WriteLine(pns);
                }
            }
            sw.WriteLine();

            // find max list in cpnlist[]
            int index = cpnlist.Length, cmax=0;
            for (int i = 0; i < cplistcount; i++)
            {
                if (cpnlist[i] == null) continue;
                if (cpnlist[i].count > cmax) { cmax = cpnlist[i].count; index = i; }
            }
            if (index < cpnlist.Length)
            {

                for (int i = 0; i < cpnlist[index].count; i++)
                {
                    mid3_4[cpnlist[index].pns[i]] = 1;
                }
            }
            Common.printArray("MID3_4 :(LAST) ", mid3_4, ref sw);
            //Common.printIndex("MID3_4 INDEX Print :(LAST) ", mid3_4, ref sw);

            // make rank 
            // temp: %sp values that is already counted
            // temp2 :  sorted array of temp as an order of ascent
            Common.printArray("MID3_4 :(temp) ", temp, ref sw);
            Common.printArray("MID3_4 :(temp2) ", temp2, ref sw);

            rank3_4 = new int[rd.Cols];
            Array.Clear(rank3_4, 0, rank3_4.Length);
            int rank = 1, prank=1;
            int idx = 0, val;
            while (temp2[temp2.Length - 1 - idx] > 0)
            {
                val = temp2[temp2.Length - 1 - idx];
                if (idx > 0 && temp2[temp2.Length - 1 - idx] == temp2[temp2.Length - 1 - (idx - 1)]) { idx++; continue; }
                prank = rank;
                for (int i = 0; i < temp.Length; i++)
                {
                    if (temp[i] == val)
                    {
                        // i: index of %sp , 
                        double intmin, intmax;
                        intmax = (i + 1) * Common.GpspInterval; 
                        intmin = intmax - Common.GpspInterval;
                        for(int k=0;k<rd.Cols;k++)
                        {
                            if (rd.PerSP[k] > intmin && rd.PerSP[k] <= intmax)
                            {
                                rank3_4[k] = prank; rank++;
                            }
                        }
                    }
                }
                idx++;
            }
            Common.printArray("MID3_4 :(RANK) ", rank3_4, ref sw);

        }
        public void printToExcel(string path1, string sheetname)
        {
            ExcelWrapper ew = new ExcelWrapper();
            if (ew.Open(path1) == false) return;
            Excel.Worksheet mySheet;
            mySheet = ew.CreateSheet(sheetname);
            if (mySheet == null) { ew.Close(); return; }
            mySheet.Range["A1:HA500"].Clear();
            ew.printLine();
            ew.printListPercent(rlist);
            ew.printLine();
            ew.printArrayPercent("MID3-4 : %SP values (this means how many times does each %SP appear in the List)", temp);
            ew.printLine();
            ew.printLine("max %SP value : " + maxSpcCount.ToString());
            ew.printLine();
            ew.printArrayPercent("MID3-4 :(Choose %SPs' that have max %SP value) ", perTemp);
            ew.printLine();

            for (int i = 0; i < perTemp.Length; i++)
            {
                if (perTemp[i] > 0)
                {
                    // get interval
                    double intmin, intmax;
                    intmax = (i + 1) * Common.GpspInterval;
                    intmin = intmax - Common.GpspInterval;

                    string pns = "";
                    pns = "Interval(%SP): " + (intmin).ToString("###0.0") + "% ~ " + intmax.ToString("###0.0") + "%    , ";
                    pns += " PN's : ";
                    // get PNs between intmax and intmin
                    for (int j = 0; j < rd.Cols; j++)
                    {
                        double persp = rd.PerSP[j];
                        if (persp > intmin && persp <= intmax)
                        {
                           pns += ((j + 1).ToString() )+" ";
                        }
                    }
                    ew.printLine(pns);
                }
            }
            ew.printLine();
            ew.printArray("MID3_4: Result:", mid3_4);
            ew.printLine();
            ew.printLine();
            ew.printArray("MID3_4: RANK:", rank3_4);
            ew.printLine();
            ew.Save();
            ew.Close();
        }
    }
}
