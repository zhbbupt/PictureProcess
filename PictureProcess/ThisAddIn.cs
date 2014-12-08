using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Threading;
using Microsoft.VisualBasic.Devices;

namespace PictureProcess
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Interop.Word.Application doc;
        public  string file_Dir;
        //private int Shapes_i;
        private static int imag_id=0;
        private string imag_save_jpg;
        private string imag_save_gif;

        private void RemoveRightRightControlBars()
        {
            //Office.CommandBarControls RightCotrlBars = Application.CommandBars.FindControls(Office.MsoControlType.msoControlButton, missing, "PictureProcess", false);
            //if (RightCotrlBars != null)
            //{
            //    foreach (Office.CommandBarControl RCB in RightCotrlBars)
            //    {
            //        RCB.Delete(true);
            //    }
            //}
            //删除所有右键自定义的菜单，以防止重复打开时菜单按钮重复
            doc = Globals.ThisAddIn.Application;
            Microsoft.Office.Core.CommandBar Bars = doc.CommandBars["Text"];
            Microsoft.Office.Core.CommandBarControls BarsContrl = Bars.Controls;
            foreach (Microsoft.Office.Core.CommandBarControl temp_contrl in BarsContrl)
            {
                string t = temp_contrl.Tag;
                if (t == "PictureProcess")
                {
                    //MessageBox.Show("Word文档中右键菜单tag:" + t.ToString(), "提示", MessageBoxButtons.OK);
                    temp_contrl.Delete(true);
                }

            }
        }
        void Application_WindowBeforeRightClick(Word.Selection Selected, ref bool Cancel)
        {
            Office.CommandBarButton RightCotrlBars = (Office.CommandBarButton)Application.CommandBars.FindControl(Office.MsoControlType.msoControlButton, missing,"PictureProcess", false);
            RightCotrlBars.Enabled = false;
            RightCotrlBars.Click -= new Office._CommandBarButtonEvents_ClickEventHandler(_RightCotrlBars_Click);
            if (Selected.Range.InlineShapes.Count > 0)
            {
                RightCotrlBars.Enabled = true;
                RightCotrlBars.Click += new Office._CommandBarButtonEvents_ClickEventHandler(_RightCotrlBars_Click);
            }
        }
        private void initiate()
        {
            doc = Globals.ThisAddIn.Application;
            if (PictureRibbon.UseDefaultFileDir == true)
            {
                file_Dir = @"C:\word图片处理工作目录" + "\\" + doc.ActiveDocument.Name;
            }
            else
            {
                file_Dir = PictureRibbon.file_Dir;
            }
            if (!Directory.Exists(file_Dir) == true)
            {
                Directory.CreateDirectory(file_Dir);
            }
        }
        void _RightCotrlBars_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            initiate();
            //Word.Selection Selected = this.Application.Selection;
            if (this.Application.Selection.InlineShapes.Count != 0)
            {
                //MessageBox.Show("Word文档中选中部分" + this.Application.Selection.Range.InlineShapes.Count.ToString(), "提示", MessageBoxButtons.OK);
                Word.InlineShape inlineShape;
                Word.Range SelectedArea= this.Application.Selection.Range;
                int CoutSelection = SelectedArea.InlineShapes.Count;
                for (int i = 1; i <= CoutSelection ; i++)
                {
                    //MessageBox.Show("Word文档中选中部分" + this.Application.Selection.Range.InlineShapes.Count.ToString(), "提示", MessageBoxButtons.OK);
                    //MessageBox.Show("Word文档中选中部分" + SelectedArea.InlineShapes.Count.ToString(), "提示", MessageBoxButtons.OK);
                    //Shapes_i = i;
                    //var shapetest = doc.Selection.this.Application.Selection.Range.InlineShapes;
                    //Word.InlineShape inlineShape = this.Application.Selection.Range.InlineShapes[3];
                    inlineShape = SelectedArea.InlineShapes[i];
                    imag_id +=1 ;
                    imag_save_jpg = file_Dir + "\\" + doc.ActiveDocument.Name.ToString() + "-" + imag_id.ToString() + "(手动选择)" + ".jpg";
                    imag_save_gif = file_Dir + "\\" + doc.ActiveDocument.Name.ToString() + "-" + imag_id.ToString() + "(手动选择)" + ".gif";
                    Thread SavePicture = new Thread(new ParameterizedThreadStart(CopyFromClipbordInlineShape));
                    SavePicture.SetApartmentState(ApartmentState.STA);
                    SavePicture.Start((object)inlineShape);
                    SavePicture.Join();
                }
                MessageBox.Show("已处理选中的" + CoutSelection.ToString() + "张图片\n保存在：" + file_Dir.ToString(), "提示", MessageBoxButtons.OK);
            }
            else
            {
                MessageBox.Show("Word文档中没有图片", "提示", MessageBoxButtons.OK);
            }
        }
        public void CopyFromClipbordInlineShape(object inlineShape_temp)
        {
            //Word.InlineShape inlineShape = this.Application.Selection.Range.InlineShapes[Shapes_i];
            Word.InlineShape inlineShape=(Word.InlineShape)inlineShape_temp;
            inlineShape.Select();
            doc.Selection.Copy();
            Computer computer = new Computer();
            if (computer.Clipboard.GetDataObject() != null)
            {
                System.Windows.Forms.IDataObject data = computer.Clipboard.GetDataObject();
                if (data.GetDataPresent(System.Windows.Forms.DataFormats.Bitmap))
                {
                    Image image = (Image)data.GetData(System.Windows.Forms.DataFormats.Bitmap, true);
                    if (PictureRibbon.imag_save_as_jpg == true)
                    {
                        image.Save(imag_save_jpg, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                    if (PictureRibbon.imag_save_as_gif == true)
                    {
                        image.Save(imag_save_gif, System.Drawing.Imaging.ImageFormat.Gif);
                    }
                }
                else
                {
                    MessageBox.Show("图片格式不正确", "提示", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("剪切板为空", "提示", MessageBoxButtons.OK);
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            RemoveRightRightControlBars();
            //添加右键按钮
            Office.CommandBarButton AddRightCotrlBar = (Office.CommandBarButton)Application.CommandBars["Text"].Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, false);
            AddRightCotrlBar.BeginGroup = true;
            AddRightCotrlBar.Tag = "PictureProcess";
            AddRightCotrlBar.Caption = "处理选中图片";
            AddRightCotrlBar.Enabled = false;
            this.Application.WindowBeforeRightClick += new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(Application_WindowBeforeRightClick);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                RemoveRightRightControlBars();
                this.Application.WindowBeforeRightClick -= new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(Application_WindowBeforeRightClick);
            }
            catch { }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
