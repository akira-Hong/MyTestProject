using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;

namespace JasonAlg
{
    public partial class CopyFile : Form
    {
        public string currentdir = null;
        public CopyFile()
        {
            InitializeComponent();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog  of= new OpenFileDialog();
            if (currentdir!=null)
                of.InitialDirectory = currentdir;
            of.Filter = "Excell files|*.xlsx;*.xls|All Files|*.*";
            if (of.ShowDialog() == DialogResult.OK)
            {
                txtCurDir.Text = of.FileName;
                FileInfo fi = new FileInfo(of.FileName);
                txtTargetFile.Text = fi.DirectoryName + "\\copy_" + fi.FullName.Substring(fi.DirectoryName.Length + 1);
            }
            else
            {
                Close();
            }
        }

        private void CopyFile_Load(object sender, EventArgs e)
        {
            btnOpen_Click( sender,  e);
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists(txtTargetFile.Text))
                {
                    MessageBox.Show("There is the same file in the directory, it can't be overwritten", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (txtCurDir.Text != txtTargetFile.Text)
                {
                    File.Copy(txtCurDir.Text, txtTargetFile.Text);
                    MessageBox.Show("Success....", "INFORMATION", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("source and target can't be the same", "ERROR",MessageBoxButtons.OK,MessageBoxIcon.Error);
                }
            }
            catch (Exception) { }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
