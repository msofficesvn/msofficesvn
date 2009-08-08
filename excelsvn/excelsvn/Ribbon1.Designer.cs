namespace excelsvn
{
    partial class Ribbon1
    {
        /// <summary>
        /// 必要なデザイナ変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナで生成されたコード

        /// <summary>
        /// デザイナ サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディタで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.Subversion = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.SVN = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.Update = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.Lock = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.Commit = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.Diff = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.Log = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.RepoBrowser = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.UnLock = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.Add = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.Delete = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.Explorer = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.File = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.Subversion.SuspendLayout();
            this.SVN.SuspendLayout();
            this.File.SuspendLayout();
            this.SuspendLayout();
            // 
            // Subversion
            // 
            this.Subversion.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Subversion.Groups.Add(this.SVN);
            this.Subversion.Groups.Add(this.File);
            this.Subversion.Label = "Subversion";
            this.Subversion.Name = "Subversion";
            // 
            // SVN
            // 
            this.SVN.Items.Add(this.Update);
            this.SVN.Items.Add(this.Lock);
            this.SVN.Items.Add(this.Commit);
            this.SVN.Items.Add(this.Diff);
            this.SVN.Items.Add(this.Log);
            this.SVN.Items.Add(this.RepoBrowser);
            this.SVN.Items.Add(this.UnLock);
            this.SVN.Items.Add(this.Add);
            this.SVN.Items.Add(this.Delete);
            this.SVN.Label = "Subversion";
            this.SVN.Name = "SVN";
            // 
            // Update
            // 
            this.Update.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Update.Label = "Update";
            this.Update.Name = "Update";
            this.Update.OfficeImageId = "FileCheckOut";
            this.Update.ShowImage = true;
            this.Update.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.Update_Click);
            // 
            // Lock
            // 
            this.Lock.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Lock.Label = "Lock";
            this.Lock.Name = "Lock";
            this.Lock.OfficeImageId = "Lock";
            this.Lock.ShowImage = true;
            this.Lock.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.Lock_Click);
            // 
            // Commit
            // 
            this.Commit.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Commit.Label = "Commit";
            this.Commit.Name = "Commit";
            this.Commit.OfficeImageId = "FileCheckIn";
            this.Commit.ShowImage = true;
            this.Commit.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.Commit_Click);
            // 
            // Diff
            // 
            this.Diff.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Diff.Label = "Diff";
            this.Diff.Name = "Diff";
            this.Diff.OfficeImageId = "ReviewCompareTwoVersions";
            this.Diff.ShowImage = true;
            this.Diff.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.Diff_Click);
            // 
            // Log
            // 
            this.Log.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Log.Label = "Log";
            this.Log.Name = "Log";
            this.Log.OfficeImageId = "ReviewTrackChanges";
            this.Log.ShowImage = true;
            this.Log.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.Log_Click);
            // 
            // RepoBrowser
            // 
            this.RepoBrowser.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RepoBrowser.Label = "RepoBrowser";
            this.RepoBrowser.Name = "RepoBrowser";
            this.RepoBrowser.OfficeImageId = "LookUp";
            this.RepoBrowser.ShowImage = true;
            this.RepoBrowser.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.RepoBrowser_Click);
            // 
            // UnLock
            // 
            this.UnLock.Label = "UnLock";
            this.UnLock.Name = "UnLock";
            this.UnLock.OfficeImageId = "AdpPrimaryKey";
            this.UnLock.ShowImage = true;
            this.UnLock.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.UnLock_Click);
            // 
            // Add
            // 
            this.Add.Label = "Add";
            this.Add.Name = "Add";
            this.Add.OfficeImageId = "OutlineExpand";
            this.Add.ShowImage = true;
            this.Add.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.Add_Click);
            // 
            // Delete
            // 
            this.Delete.Label = "Delete";
            this.Delete.Name = "Delete";
            this.Delete.OfficeImageId = "Delete";
            this.Delete.ShowImage = true;
            this.Delete.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.Delete_Click);
            // 
            // Explorer
            // 
            this.Explorer.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Explorer.Label = "Explorer";
            this.Explorer.Name = "Explorer";
            this.Explorer.OfficeImageId = "FileOpen";
            this.Explorer.ShowImage = true;
            this.Explorer.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.Explorer_Click);
            // 
            // File
            // 
            this.File.Items.Add(this.Explorer);
            this.File.Label = "File";
            this.File.Name = "File";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.Subversion);
            this.Load += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs>(this.Ribbon1_Load);
            this.Subversion.ResumeLayout(false);
            this.Subversion.PerformLayout();
            this.SVN.ResumeLayout(false);
            this.SVN.PerformLayout();
            this.File.ResumeLayout(false);
            this.File.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Subversion;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup SVN;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Update;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Commit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Lock;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Diff;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Log;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RepoBrowser;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UnLock;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Add;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Delete;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Explorer;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup File;
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
