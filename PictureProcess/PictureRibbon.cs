using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Threading;
using System.Drawing;
using System.IO;
using Microsoft.VisualBasic.Devices;
using Microsoft.Office.Core;

namespace PictureProcess
{
    public partial class PictureRibbon
    {
        private Microsoft.Office.Interop.Word.Application doc;
        private string doc_Name;
        private string work_Dir;

        public static string file_Dir;
        public static bool UseDefaultFileDir = true;

        private int Shapes_i;
        private string imag_id;
        private string imag_save_jpg;

        public static bool imag_save_as_jpg = true;

        private string imag_save_gif;

        public static bool imag_save_as_gif = false;
         
        private void PictureRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            work_Dir = @"C:\word图片处理工作目录";
            doc = Globals.ThisAddIn.Application;
            doc_Name = doc.ActiveDocument.Name;
            file_Dir = work_Dir + "\\" + doc_Name;
        }
        private void DirInitUseDefaultWorkDir()
        {
            //work_Dir = @"C:\word图片处理工作目录";
            //doc = Globals.ThisAddIn.Application;
            //doc_Name = doc.ActiveDocument.Name;
            //file_Dir = work_Dir + "\\" + doc_Name;
            if (!Directory.Exists(file_Dir) == true)
            {
                Directory.CreateDirectory(file_Dir);
            }
        }
        private void DirInitUseSelectedWorkDir()
        {
            //doc = Globals.ThisAddIn.Application;
            //doc_Name = doc.ActiveDocument.Name;
            file_Dir = work_Dir +"\\"+ doc_Name;
            if (!Directory.Exists(file_Dir) == true)
            {
                Directory.CreateDirectory(file_Dir);
            }
        }

        private void WorkDir_Click(object sender, RibbonControlEventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                work_Dir = this.folderBrowserDialog1.SelectedPath;  //获取用户选中路径
                UseDefaultFileDir = false;
                DirInitUseSelectedWorkDir();
            }
        }

        private void SaveAsJPG_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Tools.Ribbon.RibbonCheckBox SaveAsJPG = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)sender;
            if (SaveAsJPG.Checked)
            {
                imag_save_as_jpg = true;
            }
            else
            {
                imag_save_as_jpg = false;
            }
        }

        private void SaveAsGIF_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Tools.Ribbon.RibbonCheckBox SaveAsGIF = (Microsoft.Office.Tools.Ribbon.RibbonCheckBox)sender;
            if (SaveAsGIF.Checked)
            {
                imag_save_as_gif = true;
            }
            else
            {
                imag_save_as_gif = false;
            }
        }


        private void HandleALLPicture_Click(object sender, RibbonControlEventArgs e)
        {
            if (UseDefaultFileDir == true)
            {
                DirInitUseDefaultWorkDir();
            }
            if (doc.ActiveDocument.InlineShapes.Count != 0)
            {
                for (int i = 1; i <= doc.ActiveDocument.InlineShapes.Count; i++)
                {
                    Shapes_i = i;
                    imag_id = i.ToString();
                    imag_save_jpg = file_Dir + "\\" + doc.ActiveDocument.Name.ToString() + "-" + imag_id + ".jpg";
                    imag_save_gif = file_Dir + "\\" + doc.ActiveDocument.Name.ToString() + "-" + imag_id + ".gif";
                    Thread SavePicture = new Thread(CopyFromClipbordInlineShape);
                    SavePicture.SetApartmentState(ApartmentState.STA);
                    SavePicture.Start();
                    SavePicture.Join();
                }
                MessageBox.Show("Word文档中的图片已处理完毕\n保存在：" + file_Dir.ToString(), "提示", MessageBoxButtons.OK);
            }
            else
            {
                MessageBox.Show("Word文档中没有图片", "提示", MessageBoxButtons.OK);
            }
        }
        public void CopyFromClipbordInlineShape()
        {
            //var shapetest = doc.ActiveDocument.InlineShapes;
            InlineShape inlineShape = doc.ActiveDocument.InlineShapes[Shapes_i];
            inlineShape.Select();
            doc.Selection.Copy();
            Computer computer = new Computer();
            if (computer.Clipboard.GetDataObject() != null)
            {
                System.Windows.Forms.IDataObject data = computer.Clipboard.GetDataObject();
                if (data.GetDataPresent(System.Windows.Forms.DataFormats.Bitmap))
                {
                    Image image = (Image)data.GetData(System.Windows.Forms.DataFormats.Bitmap, true);
                    if (imag_save_as_jpg == true)
                    {
                        image.Save(imag_save_jpg, System.Drawing.Imaging.ImageFormat.Jpeg);
                    }
                    if (imag_save_as_gif == true)
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

        private void ClearAll_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DirectoryInfo ClearAll = new DirectoryInfo(work_Dir);
                ClearAll.Delete(true);
                Directory.CreateDirectory(work_Dir);
            }
            catch
            {
                MessageBox.Show("文件夹未创建", "提示", MessageBoxButtons.OK);
            }
        }

        private void ClearActiveDoc_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DirectoryInfo ClearActiveDoc = new DirectoryInfo(file_Dir);
                ClearActiveDoc.Delete(true);
                Directory.CreateDirectory(file_Dir);
            }
            catch
            {
                MessageBox.Show("文件夹未创建", "提示", MessageBoxButtons.OK);
            }
        }

    }
}
