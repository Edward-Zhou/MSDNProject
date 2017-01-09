using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.Threading;

namespace PowerPointAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void NewSlide_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.CustomLayout ppt_layout = Globals.ThisAddIn.Application.ActivePresentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutBlank];
            PowerPoint.Slide slide;
            slide=Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(1, ppt_layout);
            slide.Shapes[1].Delete();
            slide.Shapes.Placeholders[1].Delete();
            //slide.ApplyTheme(@"C:\Users\Karthik\Desktop\custom.thmx");
        }

        private void ExportSlide_Click(object sender, RibbonControlEventArgs e)
        {
            SlideRange sr = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange;
            //int i = 1;
            //foreach (Slide s in sr)
            //{
            //    //try
            //    //{
            //        string filePath = @"D:\" + i.ToString() + ".ppt";
            //        string fileFilter = "ppt";
            //        s.Export(filePath, fileFilter);
            //        System.Threading.Thread.Sleep(10000);
            //        //s.PublishSlides(@"D:\");
            //        i++;
            //    //}
            //    //catch (Exception ee)
            //    //{
            //    //    MessageBox.Show(e.Message);
            //    //}
            //}
            foreach (Slide s in sr)
            {
                System.Threading.Thread th = new System.Threading.Thread(() =>
                {
                    //string filePath = @"D:\" +s.SlideIndex.ToString() + ".ppt";
                    string fileFilter = "ppt";
                    
                    s.Export(@"D:\1.ppt", fileFilter);
                });
                th.SetApartmentState(ApartmentState.STA);
                th.Start();   
            }
            

            //for (int i = 1; i <= sr.Count; i++)
            //{
            //    System.Threading.Thread th = new System.Threading.Thread(() => {
            //        string filePath = @"D:\" + i.ToString() + ".ppt";
            //        string fileFilter = "ppt";
            //        sr[i].Export(filePath, fileFilter);
            //    });
            //    th.SetApartmentState(ApartmentState.STA);
            //    th.Start();    
            //}
                MessageBox.Show("ok");
        }

        private void addTaskPane_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.AddAllTaskPanes();
        }

        private void txtRange_Click(object sender, RibbonControlEventArgs e)
        {
          //MessageBox.Show( Globals.ThisAddIn.Application.ActivePresentation.Slides[1].Shapes[1].TextFrame.TextRange.Lines().Words(1).Font.Size.ToString());
            MessageBox.Show(Globals.ThisAddIn.Application.ActivePresentation.Slides[1].Shapes[1].TextFrame.TextRange.Font.Size.ToString());
        }

        private void TaskPaneWindows_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.AddAllTaskPanesWindows();
        }

        private void InsertImg_Click(object sender, RibbonControlEventArgs e)
        {
            Chart oChart;
            foreach (Slide s in Globals.ThisAddIn.Application.ActivePresentation.Slides)
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in s.Shapes)
                {
                    if (shape.HasChart==MsoTriState.msoTrue)
                    {
                        shape.Chart.Shapes.AddPicture(@"C:\Users\v-tazho\Desktop\QQ.png", MsoTriState.msoTrue, MsoTriState.msoTrue, 1, 1, 100, 50);
                    }
                }
            }
        }

        private void Quit_Click(object sender, RibbonControlEventArgs e)
        {
            //Globals.ThisAddIn.Application.Quit();
            //Microsoft.Office.Interop.PowerPoint.Presentation ppt = Globals.ThisAddIn.Application.ActivePresentation;
            //ppt.Close();
            Microsoft.Office.Interop.PowerPoint.Application ppta = Globals.ThisAddIn.Application;
            ppta.Quit();
            releaseCOM(ppta);
        }
        private static void releaseCOM(object o)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(o);
            }
            catch { }
            finally
            {
                o = null;
            }
        }

        private void Showbtn_Click(object sender, RibbonControlEventArgs e)
        {
            frmPPT f = new frmPPT();
            f.Show();
        }

        private void ShowDialogbtn_Click(object sender, RibbonControlEventArgs e)
        {
            frmPPT f = new frmPPT();
            f.ShowDialog();
        }

        private void GetSelection_Click(object sender, RibbonControlEventArgs e)
        {
            int Arg1=0;
            int Arg2=0;
            int id=0;
            Chart selection = Globals.ThisAddIn.Application.ActiveWindow.Selection as Chart;

            selection.GetChartElement(50,
                50,ref id,ref Arg1,ref Arg2);
            //selection.ShapeRange.Fill.Visible = MsoTriState.msoTrue;
            //selection.ShapeRange.Fill.ForeColor.RGB = (int)0xFF0000; 

        }
    }
}
