using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Text;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace JasonAlg
{
    public partial class Form1 : Form
    {
        public bool GOSTOP = false;
        public int step = 0;
        public RawData rd = null;
        public Mid_1 md1 = null;
        public Mid_2 md2 = null;
        public Mid_3_1 md3_1 = null;
        public Mid_3_2 md3_2 = null;
        public Mid_3_3 md3_3 = null;
        public Mid_3_4 md3_4 = null;
        public Mid_4_1 md4_1 = null;
        public Mid_4_2 md4_2 = null;
        public Mid_4_3_new1 md4_3 = null;
        public ResultDB resDB = null;
        private string MyPATH, ConfigPath;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                string programfiles = Environment.GetEnvironmentVariable("USERPROFILE");
                if (programfiles.Length < 1)
                {
                    programfiles = Environment.GetEnvironmentVariable("TMP");
                    if (programfiles.Length < 1)
                        programfiles = Directory.GetCurrentDirectory();
                }
                MyPATH = programfiles + "\\JasonAlg";
                ConfigPath = MyPATH + "\\currentfolder.cfg";
                if (!Directory.Exists(MyPATH))
                {
                    Directory.CreateDirectory(MyPATH);
                }
                GOSTOP = false;
                //if (false)
                if (File.Exists(ConfigPath))
                {
                    try
                    {
                        StreamReader sr = new StreamReader(ConfigPath);
                        txtCurDir.Text = sr.ReadLine();
                        sr.Close();
                        // make file list of Excel.
                        btnReload_Click(sender, e);
                    }
                    catch (Exception)
                    { }
                }
                else
                {
                   /*
                    IDictionary	environmentVariables = Environment.GetEnvironmentVariables();
                    foreach (DictionaryEntry de in environmentVariables)
                        {
                        Console.WriteLine("  {0} = {1}", de.Key, de.Value);
                        }
                     */
                    txtCurDir.Text = Environment.GetEnvironmentVariable("USERPROFILE");
                    if (txtCurDir.Text == "")
                        txtCurDir.Text = Directory.GetCurrentDirectory();

                    /*
                    try
                    {
                        FileAttributes attri = File.GetAttributes(ConfigPath);
                        attri = RemoveAttribute(attri, FileAttributes.System | FileAttributes.Hidden);
                        File.SetAttributes(ConfigPath, attri);
                    }
                    catch (Exception)
                    { }
                    */
                    StreamWriter sr = new StreamWriter(ConfigPath);
                    sr.WriteLine(txtCurDir.Text);
                    sr.Close();
                }
                rd = null;
                resDB = null;

                // lock 기능.'
                //File.SetAttributes(ConfigPath, FileAttributes.Hidden | FileAttributes.System);
                cmbJoblist.Enabled = false;
                btnReload.Enabled = false;
                btnCalc.Enabled = false;
                statusLabel1.Text = "먼저 Get Permission을 눌러,프로그램 실행 Permission 획득해야 합니다.";
                Common.debugpath = txtCurDir.Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "CRITICAL ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            cmbPercent.SelectedIndex = 0;
            Common.GpspInterval = 2;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            folderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer;
            if (File.Exists(ConfigPath))
            {
                try
                {
                    StreamReader sr = new StreamReader(ConfigPath);
                    folderBrowserDialog1.SelectedPath = sr.ReadLine();
                    sr.Close();
                }
                catch (Exception)
                { }
            }
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileAttributes attri = File.GetAttributes(ConfigPath);
                    attri = RemoveAttribute(attri, FileAttributes.System|FileAttributes.Hidden);
                    File.SetAttributes(ConfigPath, attri);
                    StreamWriter sr = new StreamWriter(ConfigPath);
                    sr.WriteLine(folderBrowserDialog1.SelectedPath);
                    sr.Close();
                    Directory.SetCurrentDirectory(folderBrowserDialog1.SelectedPath);
                }
                catch (Exception ex) 
                {
                    Console.WriteLine(ex.Message);
                }

                //File.SetAttributes(ConfigPath, FileAttributes.Hidden | FileAttributes.System);
                txtCurDir.Text = folderBrowserDialog1.SelectedPath;
                Common.debugpath = txtCurDir.Text;

                // make file list of Excel.
                btnReload_Click(sender, e);

            }
        }

        private void exitToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnReload_Click(object sender, EventArgs e)
        {
            // make file list of Excel.
            try
            {
                string[] filePaths = Directory.GetFiles(txtCurDir.Text, "*.xlsx",
                                             SearchOption.AllDirectories);

                cmbJoblist.Items.Clear();
                foreach (string h in filePaths)
                {
                    cmbJoblist.Items.Add(h.Substring(txtCurDir.Text.Length+1));
                }
                string[] filePaths2 = Directory.GetFiles(txtCurDir.Text, "*.xls",
                                             SearchOption.AllDirectories);
                foreach (string h in filePaths2)
                {
                    cmbJoblist.Items.Add(h.Substring(txtCurDir.Text.Length + 1));
                }
                cmbJoblist.Text = "Choose One";
                btnCalc.Enabled = false;
            }
            catch (Exception) { }
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyFile cf = new CopyFile();
            cf.currentdir = txtCurDir.Text;
            cf.ShowDialog();
        }

        private void cmbJoblist_SelectedIndexChanged(object sender, EventArgs e)
        {
            Excel.Workbook excelWorkbook = null;
            rd = null;
            resDB = null;
            GOSTOP = false;
            btnCalc.Enabled = false;
            try
            {
                txtTpn.Text = "";
                txtTpsn.Text = "";
                txtPosn.Text = "";
                string str = txtCurDir.Text+"\\"+  cmbJoblist.SelectedItem.ToString();
                if (!File.Exists(str))
                {
                    MessageBox.Show("File is not found, "+str, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Read File and display the file header.
                Excel.Application excelApp = new Excel.Application();  // Creates a new Excel Application
                excelApp.Visible = false;  // Makes Excel visible to the user.

                // The following code opens an existing workbook
                string workbookPath = str;  // Add your own path here

                try
                {
                    excelWorkbook = excelApp.Workbooks.Open(workbookPath, 0,
                        false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true,
                        false, 0, true, false, false);
                }
                catch(Exception ex)
                {
                    MessageBox.Show("File open Error:  " + ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // The following gets the Worksheets collection
                Excel.Sheets excelSheets = excelWorkbook.Worksheets;

                // The following gets Sheet1 for editing
                string currentSheet = "RawDB";
                Excel.Worksheet excelWorksheet = null;
                try
                {
                    excelWorksheet = (Excel.Worksheet)excelSheets.get_Item("RawDB");
                }
                catch
                {

                    try
                    {
                        excelWorksheet = (Excel.Worksheet)excelSheets.get_Item("Raw DB");
                    }
                    catch
                    {
                        MessageBox.Show("Error: There is not the Sheet that is named 'Raw DB'", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (excelWorkbook != null) excelWorkbook.Close();
                        return;
                    } 
                }

                // The following gets cell A1 for editing
                int tpn, tpsn, posn, RawStart, ColStart;
                try
                {
                     tpn = (int)excelWorksheet.Cells[1, 2].Value2;
                    Common.TPN = tpn;
                    txtTpn.Text = tpn.ToString();
                     tpsn = (int)excelWorksheet.Cells[2, 2].Value2;
                    Common.TSPN = tpsn;
                    txtTpsn.Text = tpsn.ToString();
                     posn = (int)excelWorksheet.Cells[3, 2].Value2;
                    Common.POSN = posn;
                    txtPosn.Text = posn.ToString();
                     RawStart = (int)excelWorksheet.Cells[4, 2].Value2;
                     ColStart = (int)excelWorksheet.Cells[4, 3].Value2;
                    txtStart.Text = "Row:" + RawStart.ToString() + ",Col:" + ColStart.ToString();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Data Read Error, please check the input data(TPN,TPSN,POSN,Start Position) in 'RAW DB' sheet.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (excelWorkbook != null) excelWorkbook.Close();
                    return;
                }

                // Read All Data and make a RawData Class
                rd = new RawData(posn,tpn, tpsn);
                int d = 0;
                for (int i = 0; i < posn; i++)
                {
                    for (int j = 0; j < tpn; j++)
                        rd.Set(i, j, 0);
                    for (int j = 0; j < tpn; j++)
                    {
                        try
                        {
                            d = Convert.ToInt32(excelWorksheet.Cells[RawStart + i, ColStart + j].Value2);
                            if (d > 0 && d <= tpn) 
                                rd.Set(i, d - 1, 1);
                        }
                        catch
                        {
                            d = 0;
                        }
                        if (d == 0) break;
                    }
                    if (i % 100 == 0)
                        Console.WriteLine("line: " + i.ToString());
                }
                excelWorkbook.Close();
                excelWorkbook = null;
                //rd.printRawData();
                btnCalc.Enabled = true;
                StatusProgress.Value = 0;
                statusLabel1.Text = "";
                timer1.Enabled = false;
            }
            catch (Exception ex) 
            {
                MessageBox.Show("Error:  unhandled error was raised. system message:" + ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (excelWorkbook != null) excelWorkbook.Close();
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnCalc_Click(object sender, EventArgs e)
        {
            if (GOSTOP == false)
            {
                StatusProgress.Value = 0;
                statusLabel1.Text = "Mid-1 Start...............";
                timer1.Enabled = true;
                GOSTOP = true;
                btnCalc.Text = "STOP";
                cmbJoblist.Enabled = false;
                button1.Enabled = false;
                btnReload.Enabled = false;
                step = 0;
            }
            else
            {
                StatusProgress.Value = 0;
                statusLabel1.Text = "STOP";
                timer1.Enabled = false;
                GOSTOP = false;
                btnCalc.Text = "GO";
                cmbJoblist.Enabled = true;
                button1.Enabled = true;
                btnReload.Enabled = true;
                step = 0;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (step == 0)  //Mid-1
            {
                timer1.Enabled = false;

                // make SPC data.
                rd.makeSPC_PerSP_APerSP();
                rd.printSPC_PerSP_APerSP();

                // make Mid-1
                md1 = new Mid_1(ref rd);
                md1.printMid1();
                string str = txtCurDir.Text + "\\" + cmbJoblist.SelectedItem.ToString();
                if (File.Exists(str))
                {
                    md1.printToExcel(str,"Mid-1");
                }
                md1.Close();

                timer1.Enabled = true;
                GC.Collect();
                StatusProgress.Value = 1*100/10;
                statusLabel1.Text = "Mid-2 Start...............";
            }
            else if (step == 1)  //Mid-2
            {
                timer1.Enabled = false;
                md2 = new Mid_2(ref rd, ref md1);
                md2.Run();
                string str = txtCurDir.Text + "\\" + cmbJoblist.SelectedItem.ToString();
                if (File.Exists(str))
                {
                    md2.printToExcel(str, "Mid-2");
                }
                md2.Close();
                timer1.Enabled = true;
                GC.Collect();
                StatusProgress.Value = 2 * 100 / 10;
                statusLabel1.Text = "Mid3-1 Start...............";
            }
            else if (step == 2) //Mid3-1
            {
                timer1.Enabled = false;
                md3_1 = new Mid_3_1(ref rd);
                string str = txtCurDir.Text + "\\" + cmbJoblist.SelectedItem.ToString();
                if (File.Exists(str))
                {
                    md3_1.printToExcel(str, "Mid3-1");
                }
                md3_1.Close();
                timer1.Enabled = true;
                GC.Collect();
                StatusProgress.Value = 3 * 100 / 10;
                statusLabel1.Text = "Mid3-2 Start...............";
            }
            else if (step == 3)//Mid3-2
            {
                timer1.Enabled = false;
                md3_2 = new Mid_3_2(ref rd);
                string str = txtCurDir.Text + "\\" + cmbJoblist.SelectedItem.ToString();
                if (File.Exists(str))
                {
                    md3_2.printToExcel(str, "Mid3-2");
                }
                md3_2.Close();
                timer1.Enabled = true;
                GC.Collect();
                StatusProgress.Value = 4* 100 / 10;
                statusLabel1.Text = "Mid3-3 Start...............";
            }
            else if (step == 4) //Mid3-3
            {
                timer1.Enabled = false;
                md3_3 = new Mid_3_3(ref rd);
                string str = txtCurDir.Text + "\\" + cmbJoblist.SelectedItem.ToString();
                if (File.Exists(str))
                {
                    md3_3.printToExcel(str, "Mid3-3");
                }
                md3_3.Close();
                timer1.Enabled = true;
                GC.Collect();
                StatusProgress.Value = 5 * 100 / 10;
                statusLabel1.Text = "Mid3-4 Start...............";
            }
            else if (step == 5) //Mid3-4
            {
                timer1.Enabled = false;
                md3_4 = new Mid_3_4(ref md3_3);
                string str = txtCurDir.Text + "\\" + cmbJoblist.SelectedItem.ToString();
                if (File.Exists(str))
                {
                    md3_4.printToExcel(str, "Mid3-4");
                }
                md3_4.Close();
                timer1.Enabled = true;
                GC.Collect();
                StatusProgress.Value = 6 * 100 / 10;
                statusLabel1.Text = "Mid4-1 Start...............";
            }
            else if (step == 6)//Mid4-1
            {
                timer1.Enabled = false;
                md4_1 = new Mid_4_1(ref rd);
                string str = txtCurDir.Text + "\\" + cmbJoblist.SelectedItem.ToString();
                if (File.Exists(str))
                {
                    md4_1.printToExcel(str, "Mid4-1");
                }
                md4_1.Close();
                timer1.Enabled = true;
                GC.Collect();
                StatusProgress.Value = 7 * 100 / 10;
                statusLabel1.Text = "Mid4-2 Start...............";
            }
            else if (step == 7)//Mid4-2
            {
                timer1.Enabled = false;
                md4_2 = new Mid_4_2(ref rd, ref md4_1);
                string str = txtCurDir.Text + "\\" + cmbJoblist.SelectedItem.ToString();
                if (File.Exists(str))
                {
                    md4_2.printToExcel(str, "Mid4-2");
                }
                md4_2.Close();
                timer1.Enabled = true;
                GC.Collect();
                StatusProgress.Value = 8 * 100 / 10;
                statusLabel1.Text = "Mid4-3 Start...............";
            }
            else if (step == 8)//Mid4-3
            {
                timer1.Enabled = false;
                md4_3 = new Mid_4_3_new1(ref rd, ref md4_1, ref md4_2);
                string str = txtCurDir.Text + "\\" + cmbJoblist.SelectedItem.ToString();
                if (File.Exists(str))
                {
                    md4_3.printToExcel(str, "Mid4-3");
                }
                md4_3.Close();
                timer1.Enabled = true;
                GC.Collect();
                StatusProgress.Value = 9 * 100 / 10;
                statusLabel1.Text = "ResultDB Start...............";
            }
            else if (step == 9)//Result DB
            {
                timer1.Enabled = false;
                resDB = new ResultDB(this);
                string str = txtCurDir.Text + "\\" + cmbJoblist.SelectedItem.ToString();
                if (File.Exists(str))
                {
                    resDB.printToExcel(str, "Result DB");
                }
                resDB.Close();
                timer1.Enabled = true;
                GC.Collect();
                StatusProgress.Value = 10 * 100 / 10;
            }
            else if (step == 10)
            {
                StatusProgress.Value =100;
                statusLabel1.Text = "Completion";
                timer1.Enabled = false;
                GOSTOP = false;
                btnCalc.Text = "GO";
                cmbJoblist.Enabled = true;
                button1.Enabled = true;
                btnReload.Enabled = true;
                step = 0;
                GC.Collect();
            }
            step++;
        }
        public string RegRead(string KeyName)
        {
            RegistryKey baseRegistryKey = Registry.LocalMachine;
            string subKey = "SOFTWARE\\JasonAlg";
            // Opening the registry key
            RegistryKey rk = baseRegistryKey;
            // Open a subKey as read-only
            RegistryKey sk1 = rk.OpenSubKey(subKey);
            // If the RegistrySubKey doesn't exist -> (null)
            if (sk1 == null)
            {
                return "";
            }
            else
            {
                try
                {
                    // If the RegistryKey exists I get its value
                    // or null is returned.
                    return (string)sk1.GetValue(KeyName.ToUpper());
                }
                catch (Exception e)
                {
                    // AAAAAAAAAAARGH, an error!
                    return "";
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            /*
            IDictionary environmentVariables = Environment.GetEnvironmentVariables();
            foreach (DictionaryEntry de in environmentVariables)
            {
                Console.WriteLine("  {0} = {1}", de.Key, de.Value);
            }
            Console.WriteLine(" ProgramFiles= {0}", Environment.GetEnvironmentVariable("ProgramFiles"));
            */
            try
            {
                string ret = postprocess.ProcessPost("req=JasonAlg");
                if (ret.Length > 5)
                {
                    string[] pp = ret.Split(new char[] { '&' });
                    Console.WriteLine(pp[0]);
                    if (pp[0] == "HELLOJASON!!!")
                    {
                        statusLabel1.Text = "성공하였습니다. Job을 선택하세요. ";
                        cmbJoblist.Enabled = true;
                        btnReload.Enabled = true;
                        button2.Enabled = false;
                        MessageBox.Show("성공하였습니다. Job을 선택하세요. ", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        statusLabel1.Text = "이런 Permission을 얻지 못했네요. www.computerlangs.com으로 접속후, 연락주세요. ";
                        MessageBox.Show("이런 Permission을 얻지 못했네요. www.computerlangs.com으로 접속후, 연락주세요. ", "No Permission", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                    statusLabel1.Text = "인터넷 연결에러입니다. ";

            }
            catch (Exception ex)
            {
                statusLabel1.Text = "인터넷 연결에러입니다.  "+ex.Message;
            }
        }

        private FileAttributes RemoveAttribute(FileAttributes attributes, FileAttributes attributesToRemove)
        {
            return attributes & ~attributesToRemove;
        }

        private void cmbPercent_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbPercent.SelectedIndex == 0) Common.GpspInterval = 2;
            else if (cmbPercent.SelectedIndex == 1) Common.GpspInterval = 1;
            if (cmbPercent.SelectedIndex == 2) Common.GpspInterval = 0.5;
        }

    }
}
