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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Diagnostics;

namespace WordPDFConvertor
{
    public partial class Form1 : Form
    {
        private List<FileToConvert> filesToConvert = new List<FileToConvert>();

        public Form1()
        {
            InitializeComponent();
        }

        public void AddFiles(string[] files)
        {
            files = files.Where(f => f.EndsWith(".docx") || f.EndsWith(".doc")).ToArray();

            filesToConvert.AddRange(files.Select(f => new FileToConvert(f, UpdateListViewInvoker)));
            UpdateListView();
        }

        private void UpdateListViewInvoker() 
        {
            if (InvokeRequired)
            {
                Invoke((Action)(() => {
                    UpdateListView();
                }));
            }
            else
            {
                UpdateListView();
            }
        }

        private void UpdateListView()
        {   
            lvFiles.Items.Clear();
            foreach (var f in filesToConvert)
            {
                lvFiles.Items.Add(f.ToListViewItem());
            }
            
        }

        private async void bConvert_Click(object sender, EventArgs e)
        {
            var app = new Word.Application();
            
            var tasks = filesToConvert.Select(f => f.Convert(
                    app,
                    (int)nudFrom.Value,
                    (int)nudTo.Value,
                    tbBegining.Text,
                    tbEnd.Text,
                    cbAddPageNumbers.Checked
                    )); 
            await Task.WhenAll(tasks);
            app.Quit();
        }


        private void bClear_Click(object sender, EventArgs e)
        {
            filesToConvert.Clear();
            UpdateListView();
        }

        private void lvFiles_DragLeave(object sender, EventArgs e)
        {
            lvFiles.BackColor = Color.White;
        }

        private void lvFiles_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                lvFiles.BackColor = Color.DeepSkyBlue;
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void lvFiles_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            AddFiles(files);

            lvFiles.BackColor = Color.White;
        }

        private void lvFiles_DoubleClick(object sender, EventArgs e)
        {
            var location = lvFiles.SelectedItems[0].SubItems[2].Text;
            if(File.Exists(location))
            {
                string argument = "/select, \"" + location + "\"";
                Process.Start("explorer.exe", argument);
            }
        }
    }
}
