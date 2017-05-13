using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework.Forms;
using PPT = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
namespace Maestro_PC
{
    public partial class frmMain : MetroForm
    {
        private PPT.Application pptApplication;
        private PPT.Presentation presentation;
        private PPT.Slide slide;
        private PPT.Slides slides;
        private int slidescount;

        public frmMain()
        {
            InitializeComponent();
        }
        private void button_Connect_Click(object sender, EventArgs e)
        {
            try
            {
                pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPT.Application;
                presentation = pptApplication.ActivePresentation;
                slides = presentation.Slides;
                slidescount = slides.Count;
                try
                {
                    // Get selected slide object in normal view 
                    slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
                }
                catch
                {
                    // Get selected slide object in reading view 
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
            catch
            {
                MetroFramework.MetroMessageBox.Show(this, "Please check PowerPoint is running.","Maestro",MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button_Previous_Click(object sender, EventArgs e)
        {
            int slideIndex = slide.SlideIndex - 1;
            if (slideIndex >= 1)
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    pptApplication.SlideShowWindows[1].View.Previous();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
            else
            {
                MetroFramework.MetroMessageBox.Show(this, "This page is the first page.", "Maestro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button_Next_Click(object sender, EventArgs e)
        {
            int slideIndex = slide.SlideIndex + 1;
            if (slideIndex > slidescount)
            {
                MetroFramework.MetroMessageBox.Show(this, "This page is the last page.", "Maestro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    pptApplication.SlideShowWindows[1].View.Next();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
        }
        private void button_Mouse_Click(object sender, EventArgs e)
        {
            presentation.SlideShowWindow.View.LaserPointerEnabled = true;
        }
        private void frmMain_Load(object sender, EventArgs e)
        {
            this.Size = new Size(500, 500);
            MetroFramework.Controls.MetroButton button_Connect = new MetroFramework.Controls.MetroButton();
            button_Connect.Size = new Size(200, 50);
            button_Connect.Left = 150;
            button_Connect.Top = 120;
            button_Connect.Text = "Connect Powerpoint";
            button_Connect.Click += new EventHandler(button_Connect_Click);
            this.Controls.Add(button_Connect);
            

            MetroFramework.Controls.MetroButton button_Previous = new MetroFramework.Controls.MetroButton();
            button_Previous.Size = new Size(200,50);
            button_Previous.Left = 150;
            button_Previous.Top = 200;
            button_Previous.Text = "Previous Slide";
            button_Previous.Click += new EventHandler(button_Previous_Click);
            this.Controls.Add(button_Previous);

            MetroFramework.Controls.MetroButton button_Next = new MetroFramework.Controls.MetroButton();
            button_Next.Size = new Size(200, 50);
            button_Next.Left = 150;
            button_Next.Top = 280;
            button_Next.Text = "Next Slide";
            button_Next.Click += new EventHandler(button_Next_Click);
            this.Controls.Add(button_Next);

            MetroFramework.Controls.MetroButton button_Mouse = new MetroFramework.Controls.MetroButton();
            button_Mouse.Size = new Size(200, 50);
            button_Mouse.Left = 150;
            button_Mouse.Top = 360;
            button_Mouse.Text = "Mouse Pointer";
            button_Mouse.Click += new EventHandler(button_Mouse_Click);
            this.Controls.Add(button_Mouse);
        }
    }
}
