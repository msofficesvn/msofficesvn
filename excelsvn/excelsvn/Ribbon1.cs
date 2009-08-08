using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection; 


namespace excelsvn
{
    public partial class Ribbon1 : OfficeRibbon
    {
        public Ribbon1()
        {
            InitializeComponent();
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Update_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application oXL = null;
            try
            {
                oXL = new Excel.Application();
                Excel.Workbook ActiveWb = Globals.ThisAddIn.Application.ActiveWorkbook;
                string DocName = ActiveWb.Name.ToString();
                MessageBox.Show(DocName);

                string RegKeyName = @"SOFTWARE\TortoiseSVN";
                string RegValueName = "ProcPath";
                RegistryKey TsvnProcPathRegkey = Registry.LocalMachine.OpenSubKey(RegKeyName);
                string TsvnProcPath = (string)TsvnProcPathRegkey.GetValue(RegValueName);
                TsvnProcPathRegkey.Close();
                //                string TsvnCommand = TsvnProcPath + " /command:update /notempfile" + " /path:" + @"""" + @"C:\work\svnonly\svntest\book6.xls" + @"""";
                //                string TsvnCommand = "\"";
                string TsvnCommand = TsvnProcPath;
                //                TsvnCommand += "\"";
                string TsvnCommandArg = " /command:update /notempfile " + DocName;
                Process.Start(TsvnCommand, TsvnCommandArg);
            }
            catch (NullReferenceException)
            {
                return;
            }
        }

        private void Lock_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void Commit_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Workbook ActiveWb = Globals.ThisAddIn.Application.ActiveWorkbook;
                string DocName = ActiveWb.FullName.ToString();
                MessageBox.Show(DocName);

                string RegKeyName = @"SOFTWARE\TortoiseSVN";
                string RegValueName = "ProcPath";
                RegistryKey TsvnProcPathRegkey = Registry.LocalMachine.OpenSubKey(RegKeyName);
                string TsvnProcPath = (string)TsvnProcPathRegkey.GetValue(RegValueName);
                TsvnProcPathRegkey.Close();
                //                string TsvnCommand = TsvnProcPath + " /command:update /notempfile" + " /path:" + @"""" + @"C:\work\svnonly\svntest\book6.xls" + @"""";
                //                string TsvnCommand = "\"";
                string TsvnCommand = TsvnProcPath;
                //                TsvnCommand += "\"";
                string TsvnCommandArg = " /command:commit /notempfile /path:" + "\"" + DocName + "\"";
                Process.Start(TsvnCommand, TsvnCommandArg);
            }
            catch (NullReferenceException)
            {
                return;
            }
        }

        private void Diff_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void Log_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void RepoBrowser_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void UnLock_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void Add_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void Delete_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void Explorer_Click(object sender, RibbonControlEventArgs e)
        {

        }

   }



}
