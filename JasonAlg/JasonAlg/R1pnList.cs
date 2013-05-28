using System;
using System.IO;
using System.Collections.Generic;

using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace JasonAlg
{
    public class R1pnList 
    {
        public  List<R1pn> rList;
        public R1pnList()
        {
            rList = new List<R1pn>(100000);
        }
        public void Add(R1pn rp)
        {
            rList.Add(rp);
        }
        public void Clear()
        {
            rList.Clear();
        }
        public int Count {
            get { return rList.Count; } 
        }
        public void FinalProcessing_PNS(ref int[] arr)
        {
            for (int i = 0; i < arr.Length; i++) arr[i] = 0;
            foreach (R1pn r in rList)
            {
                r.FinalProcessing_PNS(ref arr);
            }
        }
        public int GetMaxCount()
        {
            int maxindex = -1;
            int val=-1,index=0;
            foreach (R1pn r in rList)
            {
                if (r.cp > val) 
                { 
                    val = r.cp; 
                    maxindex = index; 
                }
                index++;
            }
           // if (maxindex < 0) 
            //    maxindex = 0;
            return val;
        }
        public bool IsExist(ref R1pn r1)
        {
            foreach (R1pn r in rList)
            {
                if (r.isSame(ref r1) == true) return true;
            }
            return false;
        }
        public R1pnList Clone()
        {
            R1pnList r = new R1pnList();
            foreach (R1pn r1 in rList)
            {
                R1pn newr = r1.Clone();
                r.Add(newr);
            }
            return r;
        }
        public void AddAllList(Int16 val)
        {
            if (rList.Count == 0)
                rList.Add(new R1pn(10));
            foreach (R1pn r1 in rList)
            {
                r1.Add(val);
            }
        }
        public void AddAllList(Int16[] val,int start, int count)
        {
            if (rList.Count == 0)
                rList.Add(new R1pn(20));
            foreach (R1pn r1 in rList)
            {
                for(int i=0; i<count; i++)
                    r1.Add(val[start+i]);
            }
        }
        public void MergeList(R1pnList rlist1)
        {
            foreach (R1pn r1 in rlist1.rList)
            {
                Add(r1);
            }            
        }
        public void getCountEachPN(ref int[] arr)
        {
            foreach (R1pn r1 in rList)
            {
                r1.getCountEachPN(ref arr);
            }
        }
        public void print(ref StreamWriter sw,int printrange=0)
        {
            sw.WriteLine("R1pnList print()-----------start----------------");
            int index = 0;
            foreach (R1pn r1 in rList)
            {
                if (printrange>0)
                    r1.print(ref sw,index);
                else
                    r1.print(ref sw,-1);
                index++;
            }
            sw.WriteLine("R1pnList print()------------end---------------");
        }
        public void printToExcel(ExcelWrapper ew)
        {

            ew.printLine("R1pnList print()-----------start----------------");
            foreach (R1pn r1 in rList)
            {
                r1.printToExcel(ew);
            }
            ew.printLine("R1pnList print()------------end---------------");
        }
        public void printPercent(ref StreamWriter sw)
        {
            sw.WriteLine("R1pnList print()-----------start----------------");
            foreach (R1pn r1 in rList)
            {
                r1.printPercent(ref sw);
            }
            sw.WriteLine("R1pnList print()------------end---------------");
        }
        public void printPercentToExcel(ExcelWrapper ew)
        {
            
            ew.printLine("R1pnList print()-----------start----------------");
            foreach (R1pn r1 in rList)
            {
                r1.printPercentToExcel(ew);
            }
            ew.printLine("R1pnList print()------------end---------------");
        }
        public void makeRankArray(ref int[] rank)
        {
            int i = 0;
            int[] temp = new int[rList.Count];
            int[] tempRank = new int[rList.Count];
            int[] realRank = new int[rList.Count];
            Array.Clear(temp, 0, temp.Length);
            Array.Clear(realRank, 0, realRank.Length);
            Array.Clear(rank, 0, rank.Length);
            i = 0;
            foreach (R1pn r1 in rList) temp[i++] = r1.cp;
            Common.makeRankArray(ref temp, ref tempRank);
            // search maxrank
            int max=0;
            for (i = 0; i < tempRank.Length; i++) if (max < tempRank[i]) max = tempRank[i];
            // get start number of each list
            int r=1, real=1,realtemp;
            while (r <= max)
            {
                realtemp=real;
                for (i = 0; i < tempRank.Length; i++)
                {
                    if (tempRank[i] == r) { realRank[i] = realtemp; real = real + temp[i]; }
                }
                r++;
            }

            i = 0;
            foreach (R1pn r1 in rList) r1.setrankarray(ref rank,realRank[i++]);
        }
    }
}
