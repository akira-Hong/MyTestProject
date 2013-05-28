using System;
using System.IO;
using System.Collections.Generic;

using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace JasonAlg
{
    public class Mid_4_3_new1
    {
        private const double variation = 0.3;  //30%
        private RawData rd;
        private Mid_4_2 md42;
        private int[] MXPlusPN, MXMinusPN;
        private double[] MXPlus, MXMinus;
        private double[] MXPlusDiff, MXMinusDiff;
        private int MXPlusCount, MXMinusCount;
        private double MXavg, MNAvg;

        //private double lower1, upper1, lower2, upper2;

        private int MXselectCount, MNselectCount;
        private int [] MXselectPN, MNselectPN;

        private StreamWriter sw;
        public int[] mid4_3;
        public Mid_4_3_new1(ref RawData r, ref Mid_4_1 md41, ref Mid_4_2 m42)
        {
            rd = r;
            md42 = m42;
            mid4_3 = new int[rd.Cols];
            Array.Clear(mid4_3, 0, mid4_3.Length);
            sw = new StreamWriter(Common.debugpath + "\\" + "MID4_3.txt");
            MXPlus = new double[rd.Cols];
            MXMinus = new double[rd.Cols];
            MXPlusPN = new int[rd.Cols];
            MXMinusPN = new int[rd.Cols];
            MXPlusCount = 0; MXMinusCount = 0;
            MXavg = 0; MNAvg = 0;
            for (int i = 0; i < md41.mx_mn.Length; i++)
            {
                if (md41.mx_mn[i] > 0) { MXPlusPN[MXPlusCount] = i; MXPlus[MXPlusCount++] = md41.mx_mn[i]; MXavg += md41.mx_mn[i]; }
                else if (md41.mx_mn[i] < 0) { MXMinusPN[MXMinusCount] = i; MXMinus[MXMinusCount++] = Math.Abs(md41.mx_mn[i]); MNAvg += Math.Abs(md41.mx_mn[i]); }
            }
            MXavg = MXavg / MXPlusCount;
            MNAvg = MNAvg / MXMinusCount;

            MXselectCount = 0; MNselectCount = 0;
            MXselectPN = new int[MXPlusCount];
            MNselectPN = new int[MXMinusCount];
            MXPlusDiff = new double[MXPlusCount];
            MXMinusDiff = new double[MXMinusCount];

            for (int i = 0; i < MXPlusCount; i++) MXPlusDiff[i] = Math.Abs(MXPlus[i] - MXavg);
            for (int i = 0; i < MXMinusCount; i++) MXMinusDiff[i] = Math.Abs(MXMinus[i] - MNAvg);
            double min1 = Double.MaxValue, min2 = Double.MaxValue;
            for (int i = 0; i < MXPlusCount; i++) 
                if (MXPlusDiff[i] < min1)
                    min1 = MXPlusDiff[i];
            for (int i = 0; i < MXMinusCount; i++)
                if (MXMinusDiff[i] < min2)
                    min2 = MXMinusDiff[i];
            for (int i = 0; i < MXPlusCount; i++)
                if (MXPlusDiff[i] == min1)
                {
                    MXselectPN[MXselectCount] = MXPlusPN[i]; MXselectCount++;
                    mid4_3[MXPlusPN[i]] = 1;
                }
            for (int i = 0; i < MXMinusCount; i++)
                if (MXMinusDiff[i] == min2)
                {
                    MNselectPN[MNselectCount] = MXMinusPN[i]; MNselectCount++;
                    mid4_3[MXMinusPN[i]] = 1;
                }
            /*
            // MX%SP
            lower1 = MXavg * (1 - variation);
            upper1 = MXavg * (1 + variation);
            for (int i = 0; i < MXPlusCount; i++)
            {
                if (MXPlus[i] >= lower1 && MXPlus[i] <= upper1)
                {
                    MXselectPN[MXselectCount] = MXPlusPN[i]; MXselectCount++;
                    mid4_3[MXPlusPN[i]] = 1;
                }
            }
            // MN%SP
            lower2 = MNAvg * (1 - variation);
            upper2 = MNAvg * (1 + variation);
            for (int i = 0; i < MXMinusCount; i++)
            {
                if (MXMinus[i] >= lower2 && MXMinus[i] <= upper2)
                {
                    MNselectPN[MNselectCount] = MXMinusPN[i]; MNselectCount++;
                    mid4_3[MXMinusPN[i]] = 1;
                }
            }
*/

            // Print Data 
            // MX List
            string one="";
            sw.WriteLine("MX PN List :  (PN, %SP) ,  Count: " + MXPlusCount.ToString());
            for (int i = 0; i < MXPlusCount; i++)
            {
                one = one + "(" + (MXPlusPN[i]+1).ToString() + ":" + MXPlus[i].ToString("#####0.00") + ") ";
                if (i > 0 && i % 5 == 0)
                {
                    sw.WriteLine(one);
                    one = "";
                }
            }
            if (one!="") sw.WriteLine(one);
            // MN List
            sw.WriteLine();
            one = "";
            sw.WriteLine("MN PN List :  (PN, %SP),  Count: " + MXMinusCount.ToString());
            for (int i = 0; i < MXMinusCount; i++)
            {
                one = one + "(" + (MXMinusPN[i]+1).ToString() + ":" + (MXMinus[i]* -1).ToString("#####0.00") + ") ";
                if (i > 0 && i % 5 == 0)
                {
                    sw.WriteLine(one);
                    one = "";
                }
            }
            if (one != "") sw.WriteLine(one);
            sw.WriteLine();
            // average
            //sw.WriteLine("AveMX%SP =" + MXavg.ToString("#######0.00") + "  Range:+/-" + (variation*100).ToString() + "% (" + lower1.ToString("######0.00") + "~" + upper1.ToString("######0.00") + ")");
            //sw.WriteLine("AveMN%SP =" + (MNAvg * -1).ToString("#######0.00") + "  Range:+/-" + (variation * 100).ToString() + "% (" + (lower2*-1).ToString("######0.00") + "~" + (upper2*-1).ToString("######0.00") + ")");
            sw.WriteLine("AveMX%SP =" + MXavg.ToString("#######0.00"));
            sw.WriteLine("AveMN%SP =" + (MNAvg * -1).ToString("#######0.00") );
            sw.WriteLine();

            // Difference table 
            one = "";
            sw.WriteLine("MX PN List :  Difference Table ,  Count: " + MXPlusCount.ToString());
            for (int i = 0; i < MXPlusCount; i++)
            {
                one = one + "(" + (MXPlusPN[i] + 1).ToString() + ":" + (MXPlus[i] - MXavg).ToString("#####0.00") + ") ";
                if (Math.Abs(MXPlus[i] - MXavg) == min1)
                    one = one + "**    ";
                if (i > 0 && i % 5 == 0)
                {
                    sw.WriteLine(one);
                    one = "";
                }
            }
            if (one != "") sw.WriteLine(one);
            // MN List
            sw.WriteLine();
            one = "";
            sw.WriteLine("MN PN List :   Difference Table ,  Count: " + MXMinusCount.ToString());
            for (int i = 0; i < MXMinusCount; i++)
            {
                one = one + "(" + (MXMinusPN[i] + 1).ToString() + ":" + (MXMinus[i] - MNAvg).ToString("#####0.00") + ") ";
                if (Math.Abs(MXMinus[i] - MNAvg)==min2)
                    one =one+"**    ";
                if (i > 0 && i % 5 == 0)
                {
                    sw.WriteLine(one);
                    one = "";
                }
            }
            if (one != "") sw.WriteLine(one);
            sw.WriteLine();

            // pn list
            sw.WriteLine();
            sw.Write("MX%SP Selected PNs' List : ");
            one = "";
            for (int i = 0; i < MXselectCount; i++)
            {
                one = one + (MXselectPN[i]+1).ToString() + " ";
            }
            sw.WriteLine(one);
            sw.Write("MN%SP Selected PNs' List : ");
            one = "";
            for (int i = 0; i < MNselectCount; i++)
            {
                one = one + (MNselectPN[i]+1).ToString() + " ";
            }
            sw.WriteLine(one);
            sw.WriteLine();

            Common.printArray("MID4_3 :MID4_3", mid4_3, ref sw);

        }
        public void Close()
        {
            sw.Close();
        }
        public void printToExcel(string path1, string sheetname)
        {
            ExcelWrapper ew = new ExcelWrapper();
            if (ew.Open(path1) == false) return;
            Excel.Worksheet mySheet;
            mySheet = ew.CreateSheet(sheetname);
            if (mySheet == null) { ew.Close(); return; }
            mySheet.Range["A1:AZ400"].Clear();
            ew.printLine();

            // Print Data 
            // MX List
            string one = "";
            ew.printLine("MX PN List :  (PN, %SP) ,  Count: " + MXPlusCount.ToString());
            for (int i = 0; i < MXPlusCount; i++)
            {
                one = one + "(" + (MXPlusPN[i] + 1).ToString() + ":" + MXPlus[i].ToString("#####0.00") + ") ";
                if (i > 0 && i % 5 == 0)
                {
                    ew.printLine(one);
                    one = "";
                }
            }
            if (one != "") ew.printLine(one);
            // MN List
            ew.printLine();
            one = "";
            ew.printLine("MN PN List :  (PN, %SP),  Count: " + MXMinusCount.ToString());
            for (int i = 0; i < MXMinusCount; i++)
            {
                one = one + "(" + (MXMinusPN[i] + 1).ToString() + ":" + (MXMinus[i] * -1).ToString("#####0.00") + ") ";
                if (i > 0 && i % 5 == 0)
                {
                    ew.printLine(one);
                    one = "";
                }
            }
            if (one != "")  ew.printLine(one);
            ew.printLine();

            // average
            //ew.printLine("AveMX%SP = " + MXavg.ToString("#######0.00") + "  Range:+/-" + (variation * 100).ToString() + "% (" + lower1.ToString("######0.00") + "~" + upper1.ToString("######0.00") + ")");
            //ew.printLine("AveMN%SP = " + (MNAvg * -1).ToString("#######0.00") + "  Range:+/-" + (variation * 100).ToString() + "% (" + (lower2 * -1).ToString("######0.00") + "~" + (upper2 * -1).ToString("######0.00") + ")");
            ew.printLine("AveMX%SP = " + MXavg.ToString("#######0.00") );
            ew.printLine("AveMN%SP = " + (MNAvg * -1).ToString("#######0.00") );
            // pn list
            ew.printLine();
            one = "MX%SP Selected PNs' List : ";
            for (int i = 0; i < MXselectCount; i++)
            {
                one = one + (MXselectPN[i] + 1).ToString() + " ";
            }
            ew.printLine(one);

            one = "MN%SP Selected PNs' List : ";
            for (int i = 0; i < MNselectCount; i++)
            {
                one = one + (MNselectPN[i] + 1).ToString() + " ";
            }
            ew.printLine(one);
            ew.printLine();

            ew.printArray("MID4_3: Result:", mid4_3);
            ew.printLine();
            ew.Save();
            ew.Close();
        }
    }
}
