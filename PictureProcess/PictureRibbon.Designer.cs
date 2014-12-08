namespace PictureProcess
{
    partial class PictureRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public PictureRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.PictureProcess = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.AllPictures = this.Factory.CreateRibbonButton();
            this.SaveAsJPG = this.Factory.CreateRibbonCheckBox();
            this.SaveAsGIF = this.Factory.CreateRibbonCheckBox();
            this.HandleALLPicture = this.Factory.CreateRibbonButton();
            this.ClearAll = this.Factory.CreateRibbonButton();
            this.ClearActiveDoc = this.Factory.CreateRibbonButton();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.PictureProcess.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // PictureProcess
            // 
            this.PictureProcess.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.PictureProcess.Groups.Add(this.group1);
            this.PictureProcess.Label = "图片处理";
            this.PictureProcess.Name = "PictureProcess";
            // 
            // group1
            // 
            this.group1.Items.Add(this.AllPictures);
            this.group1.Items.Add(this.SaveAsJPG);
            this.group1.Items.Add(this.SaveAsGIF);
            this.group1.Items.Add(this.HandleALLPicture);
            this.group1.Items.Add(this.ClearAll);
            this.group1.Items.Add(this.ClearActiveDoc);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // AllPictures
            // 
            this.AllPictures.Label = "工作目录选择";
            this.AllPictures.Name = "AllPictures";
            this.AllPictures.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WorkDir_Click);
            // 
            // SaveAsJPG
            // 
            this.SaveAsJPG.Checked = true;
            this.SaveAsJPG.Label = "保存为jpg";
            this.SaveAsJPG.Name = "SaveAsJPG";
            this.SaveAsJPG.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveAsJPG_Click);
            // 
            // SaveAsGIF
            // 
            this.SaveAsGIF.Label = "保存为GIF";
            this.SaveAsGIF.Name = "SaveAsGIF";
            this.SaveAsGIF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveAsGIF_Click);
            // 
            // HandleALLPicture
            // 
            this.HandleALLPicture.Label = "处理图片";
            this.HandleALLPicture.Name = "HandleALLPicture";
            this.HandleALLPicture.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.HandleALLPicture_Click);
            // 
            // ClearAll
            // 
            this.ClearAll.Label = "清空工作目录";
            this.ClearAll.Name = "ClearAll";
            this.ClearAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ClearAll_Click);
            // 
            // ClearActiveDoc
            // 
            this.ClearActiveDoc.Label = "清空当前目录";
            this.ClearActiveDoc.Name = "ClearActiveDoc";
            this.ClearActiveDoc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ClearActiveDoc_Click);
            // 
            // PictureRibbon
            // 
            this.Name = "PictureRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.PictureProcess);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.PictureRibbon_Load);
            this.PictureProcess.ResumeLayout(false);
            this.PictureProcess.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab PictureProcess;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AllPictures;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox SaveAsJPG;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox SaveAsGIF;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton HandleALLPicture;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ClearAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ClearActiveDoc;
    }

    partial class ThisRibbonCollection
    {
        internal PictureRibbon PictureRibbon
        {
            get { return this.GetRibbon<PictureRibbon>(); }
        }
    }
}
