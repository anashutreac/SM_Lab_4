using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;

namespace Lab_4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void converterToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.textBox1.Text = ofd.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            axWindowsMediaPlayer1.URL = textBox1.Text;
            axWindowsMediaPlayer1.Ctlcontrols.play();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            axWindowsMediaPlayer1.Ctlcontrols.stop();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            axWindowsMediaPlayer1.Ctlcontrols.pause();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string fileName = @"D:\Univer\Anul-III\Sem-II\SM\Lab4\SM_Lab_4\curs2.ppt";
            string exportName = "video_of_presentation";
            string exportPath = @"D:\Univer\Anul-III\Sem-II\SM\Lab4\SM_Lab_4\{0}.mp4";

            Microsoft.Office.Interop.PowerPoint.Application ppApp = new Microsoft.Office.Interop.PowerPoint.Application();
            ppApp.Visible = MsoTriState.msoTrue;
            ppApp.WindowState = PpWindowState.ppWindowMinimized;
            Microsoft.Office.Interop.PowerPoint.Presentations oPresSet = ppApp.Presentations;
            Microsoft.Office.Interop.PowerPoint._Presentation oPres = oPresSet.Open(fileName,
                        MsoTriState.msoFalse, MsoTriState.msoFalse,
                        MsoTriState.msoFalse);
            try
            {

                oPres.CreateVideo(exportName);
                oPres.SaveCopyAs(String.Format(exportPath, exportName),
                    PowerPoint.PpSaveAsFileType.ppSaveAsWMV,
                    MsoTriState.msoCTrue);
            }
            finally
            {
                ppApp.Quit();
            }
        }
    }
}
