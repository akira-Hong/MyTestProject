using System;
using System.IO;
using System.Collections.Generic;

using System.Text;

namespace JasonAlg
{
    public class Mid_3_3_table
    {
        public int Rows, Cols, tpsn;
        public int[,] rawdata;
        public int[] spc;
        public int spcindex, spcmax;
        public Mid_3_3_table(ref RawData r)
        {
            Rows = r.Rows; Cols = (int)(100/Common.GpspInterval); tpsn = r.tpsn;
            rawdata = new int[Rows, Cols];
            spc = new int[Cols];
            for (int i = 0; i < Rows; i++) for (int j = 0; j < Cols; j++) rawdata[i, j] = 0;

            // make MID3_3_table
            for (int i = 0; i < r.Rows; i++)
                for (int j = 0; j < r.Cols; j++)
                {
                    if (r.rawdata[i, j] > 0)
                    {
                        rawdata[i, Common.CalcIndex(r.PerSP[j])]++;
                    }
                }
            // sum each cols and get max out of spcs.
            int count;
            spcindex = 0;
            spcmax = -1;
            for (int i = 0; i < Cols; i++)
            {
                count = 0;
                for (int j = 0; j < Rows; j++)
                {
                    count += rawdata[j, i];
                }
                spc[i] = count;
                if (count > spcmax) { spcmax = count; spcindex = i; }
            }
        }
        public void printtable(ref StreamWriter sw)
        {
            sw.WriteLine("MID3_3_Table Print ------------------------------------");
            sw.Write(" PN ");
            for (int j = 0; j < Cols; j++)
            {
                sw.Write("{0,5}", ((j+1)*Common.GpspInterval).ToString("##0.0"));
            }
            sw.WriteLine();
            for (int i = 0; i < Rows; i++)
            {
                sw.Write("{0,4}", (i+1).ToString("##0"));
                for (int j = 0; j < Cols; j++)
                {
                    sw.Write("{0,5}", rawdata[i, j].ToString("####0"));
                }
                sw.WriteLine();
            }
            sw.WriteLine();
            sw.WriteLine("MID3_3_Table SPC ------------------------------------");
            for (int j = 0; j < Cols; j++)
            {
                sw.Write("{0,5}", ((j + 1) * Common.GpspInterval).ToString("##0.0"));
            }
            sw.WriteLine();
            for (int j = 0; j < Cols; j++)
            {
                sw.Write("{0,5}", spc[j].ToString("####0"));
            }
            sw.WriteLine();
            sw.WriteLine("MID3_3_Table spcIndex : " + spcindex.ToString()+ ", max PN : "+((spcindex+1)*Common.GpspInterval).ToString());
            sw.WriteLine("MID3_3_Table spcmax   : "+ spcmax.ToString());
            sw.WriteLine();
        }
    }
}
