using System;
using System.IO;
using System.Collections.Generic;

using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace JasonAlg
{
    public class R1pn
    {
        public Int16[] r1pn;
        public int cp;
        public int indexofbaseR1pn;
        private bool sorted;
        private UInt16[] r1pnSorted;
        public R1pn(int len)
        {
            r1pn = new Int16[len];
            Array.Clear(r1pn, 0, r1pn.Length);
            //r1pnSorted = new int[len]; 
            cp = 0;
            indexofbaseR1pn = 0;
            sorted = false;
            r1pnSorted = null;
        }
        public void Add(Int16 pn)
        {
            Int16[] temp;
            if (r1pn.Length > cp)
            {
                r1pn[cp] = pn;
                cp++;
            }
            else
            {
                temp = new Int16[r1pn.Length + 20];
                Array.Clear(temp, 0, temp.Length);
                Array.Copy(r1pn, temp, cp);
                r1pn = temp;
                temp = null;
                r1pn[cp] = pn;
                cp++;
            }
            sorted = false;
        }
        public void CopyFrom(R1pn r)
        {
            cp = r.cp;
            indexofbaseR1pn = r.indexofbaseR1pn;
            for (int i = 0; i < cp; i++) r1pn[i] = r.r1pn[i];
            sorted = false;
        }
        public void Clear()
        {
            cp = 0;
        }
        public void ShiftLeft_1()
        {
            if (cp <= 0) return;
            cp--;
            for (int i = 0; i < cp; i++)
                r1pn[i] = r1pn[i + 1];
            sorted = false;
        }
        public bool isSame(ref R1pn r1)
        {
            /*
            if (!sorted)
            {
                r1pnSorted = new int[cp]; 
                Array.Copy(r1pn, r1pnSorted, cp);
                Array.Sort(r1pnSorted);
                sorted = true;
            }
            if (!r1.sorted)
            {
                r1.r1pnSorted = new int[r1.cp]; 
                Array.Copy(r1.r1pn, r1.r1pnSorted, r1.cp);
                Array.Sort(r1.r1pnSorted);
                r1.sorted = true;
            }
            */
            if (cp != r1.cp) return false;
            for (int i = 0; i < cp; i++)
              //  if (r1pnSorted[i] != r1.r1pnSorted[i]) return false;
                if (r1pn[i] != r1.r1pn[i]) return false;
            if (indexofbaseR1pn != r1.indexofbaseR1pn) return false;
            return true;
        }
        public void FinalProcessing_PNS(ref int[] arr)
        {
            for (int i = 0; i < cp; i++)
                arr[r1pn[i]] = arr[r1pn[i]]+1;
        }
        public void print(ref StreamWriter sw, int index=-1)
        {
            //sw.Write(indexofbaseR1pn.ToString() + " ");
            // print Range
            double intmin, intmax;
            if (index >= 0)
            {
                intmax = (index + 1) * Common.GpspInterval;
                intmin = intmax - Common.GpspInterval;
                sw.Write((intmin).ToString("###0.0") + "% ~ " + intmax.ToString("###0.0") + "% :");
            }
            sw.Write("Count: " + cp.ToString() + " ");
            sw.Write("List: ");
            for (int i = 0; i < cp; i++) sw.Write((r1pn[i] + 1).ToString() + " ");
            sw.WriteLine("");
        }
        public void printToExcel(ExcelWrapper ew)
        {
            //sw.Write(indexofbaseR1pn.ToString() + " ");
            string line = "";
            line = "Count: " + cp.ToString() + " ";
            line += "List: ";
            for (int i = 0; i < cp; i++) line += (r1pn[i] + 1).ToString() + " ";
            ew.printLine(line);
        }
        public void printPercent(ref StreamWriter sw)
        {
            //sw.Write(indexofbaseR1pn.ToString() + " ");
            sw.Write("Count: "+cp.ToString() + " ");
            sw.Write("List: ");
            for (int i = 0; i < cp; i++) sw.Write(((r1pn[i] + 1)*Common.GpspInterval).ToString() + " ");
            sw.WriteLine("");
        }
        public void printPercentToExcel(ExcelWrapper ew)
        {
            //sw.Write(indexofbaseR1pn.ToString() + " ");
            string line = "";
            line="Count: " + cp.ToString() + " ";
            line +="List: ";
            for (int i = 0; i < cp; i++) line += ((r1pn[i] + 1) * Common.GpspInterval).ToString() + " ";
            ew.printLine(line);
        }
        public R1pn Clone()
        {
            R1pn r = new R1pn(this.cp + 10);
            Array.Copy(this.r1pn, r.r1pn, cp);
            r.cp = cp;
            r.indexofbaseR1pn = indexofbaseR1pn;
            return r;
        }
        public void MakeMultiple(int n)
        {
            int r = cp % n;
            if (r == 0) return;
            r = n - r;
            for (int i = 0; i < r; i++) Add(-1);
        }
        public void getCountEachPN(ref int[] arr)
        {
            for (int i = 0; i < cp; i++)
                if (r1pn[i] >= 0 && r1pn[i]<arr.Length) arr[r1pn[i]]++;
        }
        public string  getString()
        {
            //sw.Write(indexofbaseR1pn.ToString() + " ");
            string line = "";
            for (int i = 0; i < cp; i++) line += (r1pn[i] + 1).ToString() + " ";
            return line;
        }
        public void setrankarray(ref int[]rank,int ranknum)
        {
            for (int i = 0; i < cp; i++) rank[r1pn[i]] = ranknum;
        }
    }
}
