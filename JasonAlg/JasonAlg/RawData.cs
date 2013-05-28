using System;
using System.Collections.Generic;

using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace JasonAlg
{
    public class RawData
    {
        public int Rows, Cols, tpsn;
        public int[,] rawdata;
        public int[] spc;
        public double[] PerSP;
        public double APerSP;
        public RawData(int r, int c, int t)
        {
            Rows = r; Cols = c; tpsn = t;
            rawdata = new int[r, c];
            spc = new int[c];
            PerSP = new double[c];
        }
        public void Set(int r, int c, int d) { rawdata[r, c] = d; }
        public void printRawData()
        {
            Console.WriteLine("RawData Print ------------------------------------");
            for (int i = 0; i < Rows; i++)
            {
                for (int j = 0; j < Cols; j++)
                {
                    Console.Write(rawdata[i, j] + " ");
                }
                Console.WriteLine();
            }
        }
        public void printSPC_PerSP_APerSP()
        {
            Console.WriteLine("SPC Print ------------------------------------");
            for (int c = 0; c < Cols; c++)
            {
                Console.Write(spc[c].ToString() + " ");
            }
            Console.WriteLine();
            Console.WriteLine("%SP Print ------------------------------------");
            for (int c = 0; c < Cols; c++)
            {
                Console.Write(PerSP[c].ToString("####0.00") + " ");
            }
            Console.WriteLine();
            Console.WriteLine("A%SP : "+APerSP.ToString("#####0.00"));
        }
        public void makeSPC_PerSP_APerSP()
        {
            int count;
            for (int c = 0; c < Cols; c++)
            {
                count = 0;
                for (int r = 0; r < Rows; r++)
                {
                    if (rawdata[r, c] > 0) count++;
                }
                spc[c] = count;
            }
            double sum = 0;
            for (int c = 0; c < Cols; c++)
            {
                PerSP[c] = (double)spc[c] * 100 / Rows;
                sum += PerSP[c];
            }
            APerSP = sum / Cols;
        }
    }
}
