using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Win32;

namespace WordPDFConvertor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public void AddFiles(string[] files)
        {
            files = files.Where(f => f.EndsWith(".docx") || f.EndsWith(".doc")).ToArray();
            lbFiles.Items.AddRange(files);
        }

        private void lbFiles_DragLeave(object sender, EventArgs e)
        {
            lbFiles.BackColor = Color.White;
        }

        private void lbFiles_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                lbFiles.BackColor = Color.DeepSkyBlue;
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void lbFiles_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            AddFiles(files);

            lbFiles.BackColor = Color.White;
        }

        private void bConvert_Click(object sender, EventArgs e)
        {
            var app = new Word.Application();
            foreach (string file in lbFiles.Items)
            {
                if (File.Exists(file))
                {
                    var document = app.Documents.Open(file);

                    Word.WdStatistic PagesCountStat = Word.WdStatistic.wdStatisticPages;
                    int pagesCount = document.ComputeStatistics(PagesCountStat);
                    var (from, to) = GetRange(pagesCount);
                    string saveAs = $"{document.Path}\\{from}-{to} {document.Name}".Replace(".docx", ".pdf").Replace(".doc", ".pdf");
                    
                    document.ExportAsFixedFormat(
                        saveAs,
                        Word.WdExportFormat.wdExportFormatPDF,
                        Range:Word.WdExportRange.wdExportFromTo,
                        From:from,
                        To:to
                    );

                    document.Close();        
                }
            }
            app.Quit();
        }

        private Tuple<int, int> GetRange(int pagesCount)
        {
            int from = (int)nudFrom.Value;
            int to = (int)nudTo.Value;

            if (to == 0)
                to = pagesCount;

            if (to < 0 && to > -pagesCount)
                to = pagesCount + to;

            if (to < -pagesCount)
                to = pagesCount;

            if (from > pagesCount)
                from = 1;

            if (from >= to)
            {
                from = 1;
                to = pagesCount;
            }
            

            return new Tuple<int, int>(from, to);
        }

        private void bClear_Click(object sender, EventArgs e)
        {
            lbFiles.Items.Clear();
        }
    }
}
