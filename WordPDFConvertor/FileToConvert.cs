using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Xml.Linq;
using Task = System.Threading.Tasks.Task;
using static WordPDFConvertor.FileToConvert;

namespace WordPDFConvertor
{
    public enum Status
    {
        READY_TO_CONVERT,
        PROCESSING,
        DONE,
        FAILED
    }

    public class FileToConvert
    {
        public string PathToFile { get; private set; }
        public string SavedAs { get; private set; }
        public Status Status { get; private set; }

        public delegate void SomethingChanged();
        public event SomethingChanged OnSomethingChanged;
        public FileToConvert(string pathToFile, SomethingChanged eventHandler)
        {
            PathToFile = pathToFile;
            SavedAs = "NOT SAVED";
            Status = Status.READY_TO_CONVERT;
            OnSomethingChanged += eventHandler;
        }   

        public async Task Convert(
            Word.Application app,
            int fromRelative, 
            int toRelative,
            string beginingText,
            string endText,
            bool addPageNumbersToName)
        {
            await Task.Run(() =>
            {
                if (File.Exists(PathToFile))
                {
                    try
                    {
                        Status = Status.PROCESSING;
                        OnSomethingChanged?.Invoke();
                        var document = app.Documents.Open(PathToFile);

                        Word.WdStatistic PagesCountStat = Word.WdStatistic.wdStatisticPages;
                        int pagesCount = document.ComputeStatistics(PagesCountStat);

                        var (from, to) = GetRange(pagesCount, fromRelative, toRelative);
                        var saveAs = CreateName(document, from, to, beginingText, endText, addPageNumbersToName);

                        document.ExportAsFixedFormat(
                            saveAs,
                            Word.WdExportFormat.wdExportFormatPDF,
                            Range: Word.WdExportRange.wdExportFromTo,
                            From: from,
                            To: to
                        );

                        document.Close();
                        SavedAs = saveAs;
                        Status = Status.DONE;
                        OnSomethingChanged?.Invoke();
                    }
                    catch
                    {
                        Status = Status.FAILED;
                        OnSomethingChanged?.Invoke();
                    }
                }
                else
                {
                    Status = Status.FAILED;
                    OnSomethingChanged?.Invoke();
                }

            });
        }

        private string CreateName(Document document, 
            int fromPage, 
            int ToPage, 
            string beginingText,
            string endText,
            bool addPageNumbersToName)
        {
            

            string fileName = addPageNumbersToName ? $"{fromPage}-{ToPage} " : "" +
                (beginingText.Length > 0 ? beginingText + " " : "") +
                document.Name.Replace(".docx", "").Replace(".doc", "") +
                (endText.Length > 0 ? " " + endText : "");

            string saveAs = $"{document.Path}\\{fileName}.pdf";

            return saveAs;
        }

        private Tuple<int, int> GetRange(int pagesCount, int fromRelative, int toRelative)
        {
            if (toRelative == 0)
                toRelative = pagesCount;

            if (toRelative < 0 && toRelative > -pagesCount)
                toRelative = pagesCount + toRelative;

            if (toRelative < -pagesCount)
                toRelative = pagesCount;

            if (fromRelative > pagesCount)
                fromRelative = 1;

            if (fromRelative >= toRelative)
            {
                fromRelative = 1;
                toRelative = pagesCount;
            }


            return new Tuple<int, int>(fromRelative, toRelative);
        }

        public ListViewItem ToListViewItem()
        {
            ListViewItem item = new ListViewItem(PathToFile);
            item.SubItems.Add(Status.ToString());
            item.SubItems.Add(SavedAs);

            return item;
        }
    }
}
