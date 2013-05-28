using System;
using System.IO;
using System.Collections.Generic;

using System.Text;

namespace JasonAlg
{
    public class Mid_3_2_Table
    {
        public int Rows, Cols;
        public double[,] table;
        public double[] avg;
        public int maxindex;
        public double maxavg;
        public Mid_3_2_Table(int r, int c)
        {
            Rows = r; Cols = c;
            table = new double[r, c];
            avg = new double[c];
        }
        public void Set(int r, int c, double d) { table[r, c] = d; }
        public void makeAVERAGE()
        {
            double val = 0;
            int  maxindex = 0; ;
            maxavg = 0;
            for (int c = 0; c < Cols; c++)
            {
                val = 0;
                for (int r = 0; r < Rows; r++)
                {
                    if (table[r, c] > 0) val = val + table[r, c];
                }
                avg[c] = val / Rows;
                if (avg[c] > maxavg)
                {
                    maxindex = c;
                    maxavg = avg[c];
                }
            }
        }
        public void printTable(ref StreamWriter sw)
        {
            sw.WriteLine("MID3_2_Table Print ------------------------------------");
            for (int i = 0; i < Rows; i++)
            {
                for (int j = 0; j < Cols; j++)
                {
                    sw.Write("{0,6}", table[i, j].ToString("##0.0"));
                }
                sw.WriteLine();
            }
            sw.WriteLine();
            sw.WriteLine("MID3_2_Table AVERAGE ------------------------------------");
            for (int j = 0; j < Cols; j++)
            {
                sw.Write("{0,7}",avg[j].ToString("##0.00"));
            }
            sw.WriteLine();
           // sw.WriteLine("MID3_2_Table: MaxIndex : "+maxindex.ToString());
            sw.WriteLine("MID3_2_Table: MaxAVG : " + maxavg.ToString("####0.00"));
        }
    }
}
