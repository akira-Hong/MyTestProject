using System;
using System.IO;
using System.Collections.Generic;

using System.Text;
using System.Collections;

namespace JasonAlg
{
    public class myReverserClass : IComparer
    {
        // Calls CaseInsensitiveComparer.Compare with the parameters reversed.
        int IComparer.Compare(Object x, Object y)
        {
            return ((new CaseInsensitiveComparer()).Compare(y, x));
        }
    }
    public static class Common
    {
        public static double GpspInterval = 0.5;  // G%SP : 2% 
        public static int TPN = 0;
        public static int TSPN = 0;
        public static int POSN = 0;
        public static string debugpath;
        public static int MaxValue(ref int[] arr)
        {
            int max = Int32.MinValue;
            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i] > max) max = arr[i];
            }
            return max;
        }
        public static double MaxValue(ref double[] arr)
        {
            double max = Double.MinValue;
            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i] > max) max = arr[i];
            }
            return max;
        }
        public static int MaxIndex(ref int[] arr)
        {
            int max = Int32.MinValue;
            int index = 0;
            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i] > max) { max = arr[i]; index = i; }
            }
            return index;
        }
        public static int MaxIndex(ref double[] arr)
        {
            double max = Double.MinValue;
            int index=0;
            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i] > max) { max = arr[i]; index = i; }
            }
            return index;
        }
        public static int MinIndex(ref double[] arr)
        {
            double min = Double.MaxValue;
            int index = 0;
            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i] < min) { min = arr[i]; index = i; }
            }
            return index;
        }
        public static int MinIndexOutofPlus(ref double[] arr)
        {
            double min = Double.MaxValue;
            int index = 0;
            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i] <= 0) continue;
                if (arr[i] < min) { min = arr[i]; index = i; }
            }
            return index;
        }
        public static int MaxIndexOutofMinus(ref double[] arr)
        {
            double min = Double.MinValue;
            int index = 0;
            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i] >= 0) continue;
                if (arr[i] > min) { min = arr[i]; index = i; }
            }
            return index;
        }
        public static void printArray(string header, int[] arr, ref StreamWriter sw)
        {
            sw.WriteLine("ARRAYPRINT: " + header);
            for (int i = 0; i < arr.Length; i++)
            {
                sw.Write("{0,4}",(i+1).ToString("##0"));
            }
            sw.WriteLine();
            for (int i = 0; i < arr.Length; i++)
            {
                sw.Write("{0,4}",arr[i].ToString("##0"));
            }
            sw.WriteLine();
        }

        public static void printArray(string header, double[] arr, ref StreamWriter sw)
        {
            sw.WriteLine("ARRAYPRINT: " + header);
            for (int i = 0; i < arr.Length; i++)
            {
                sw.Write("{0,7}", (i + 1).ToString("####0"));
            }
            sw.WriteLine();
            for (int i = 0; i < arr.Length; i++)
            {
                sw.Write("{0,7}", arr[i].ToString("###0.00"));
            }
            sw.WriteLine();
        }
        public static void printArrayPercent(string header, int[] arr, ref StreamWriter sw)
        {
            sw.WriteLine("ARRAYPRINT: " + header);
            for (int i = 0; i < arr.Length; i++)
            {
                sw.Write("{0,5}", ((i + 1)*Common.GpspInterval).ToString("##0.0"));
            }
            sw.WriteLine();
            for (int i = 0; i < arr.Length; i++)
            {
                sw.Write("{0,5}", arr[i].ToString("##0"));
            }
            sw.WriteLine();
        }
        public static void printIndex(string header, int[] arr, ref StreamWriter sw)
        {
            sw.WriteLine("ARRAYINDEXPRINT: " + header);
            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i]>0)
                    sw.Write("{0,4}",(i+1).ToString("##0"));
            }
            sw.WriteLine();
        }
        public static void printIndex(string header, double[] arr, ref StreamWriter sw)
        {
            sw.WriteLine("ARRAYINDEXPRINT: " + header);
            for (int i = 0; i < arr.Length; i++)
            {
                if (arr[i] > 0)
                    sw.Write("{0,4}", (i + 1).ToString("##0"));
            }
            sw.WriteLine();
        }
        public static int CalcIndex(double persp)
        {
            int isp = (int)persp;
            if ((double)isp == persp)
            {
                persp = persp - (Common.GpspInterval / 2.0);
            }
            int index=0;
            if (persp >= 100) return ((int)(100 / Common.GpspInterval)) - 1;

            if (persp > 0) index = (int)(persp / Common.GpspInterval);
            return index;
        }
        public static void printRankArray(string header, int[] arr, ref StreamWriter sw)
        {
            int[] spc = new int[arr.Length];
            int[] rank = new int[arr.Length];
            IComparer myComparer = new myReverserClass();
            Array.Copy(arr, spc, arr.Length);
            Array.Sort(spc, myComparer);
            int prev = Int32.MinValue;
            for (int i = 0; i < arr.Length; i++)
            {
                if (prev != spc[i])
                {
                    for(int j=0;j<arr.Length;j++)
                    {
                        if (arr[j] == spc[i]) rank[j] = i + 1;
                    }
                    prev = spc[i];   
                }
            }
            printArray(header, rank, ref sw);
        }
        public static void printArray2(string header, int[] arr, ref StreamWriter sw)
        {
            sw.WriteLine("ARRAYPRINT: " + header);
            for (int i = 0; i < arr.Length; i++)
            {
                sw.Write((i + 1).ToString("##0") + ":" + arr[i].ToString("######0") +", ");
                if (i > 0 && i % 15 == 0) sw.WriteLine();
            }
            sw.WriteLine();
        }
        public static void makeRankArray(ref int[] source, ref int[] rank)
        {
            int[] spc = new int[source.Length];
            IComparer myComparer = new myReverserClass();
            Array.Copy(source, spc, source.Length);
            Array.Sort(spc, myComparer);
            int prev = Int32.MinValue;
            for (int i = 0; i < source.Length; i++)
            {
                if (prev != spc[i])
                {
                    for (int j = 0; j < source.Length; j++)
                    {
                        if (source[j] == spc[i]) rank[j] = i + 1;
                    }
                    prev = spc[i];
                }
            }
        }
    }
}
