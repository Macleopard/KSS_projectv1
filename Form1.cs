using System;

using System.IO;

using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Xceed.Words.NET;
using Xceed.Workbooks.NET;
using Application = Microsoft.Office.Interop.Word.Application;
using Excel = Microsoft.Office.Interop.Excel;

namespace KSS_project
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //  для сохранения пдф в файл
        static void SaveOutputDoc(PdfDocument outputPDFDoc, int pageNo, string outFolderPath)
        {
            string outputPDFFilePath = Path.Combine(outFolderPath, pageNo.ToString() + ".pdf");
            outputPDFDoc.Save(outputPDFFilePath);
        }

        static void splitPDF(string inFolderPath, string inFileName, string outputPath)
        {
            string inFilePath = Path.Combine(inFolderPath, inFileName);
            PdfDocument inFile = PdfReader.Open(inFilePath, PdfDocumentOpenMode.Import);
            var totalPagesInInpFile = inFile.PageCount;
            while (totalPagesInInpFile != 0)
            {
                PdfDocument outputPDFDoc = new PdfDocument();
                outputPDFDoc.AddPage(inFile.Pages[totalPagesInInpFile - 1]);
                SaveOutputDoc(outputPDFDoc, totalPagesInInpFile, outputPath);
                totalPagesInInpFile--;
            }

            MessageBox.Show("splitting pdf completed");
        }

        private void w2pdf_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "word files (*.docx)|*.docx|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var fileName = openFileDialog.SafeFileName; // только имя файла без пути
                    var path = Path.GetDirectoryName(openFileDialog.FileName);
                    var pdfName = "ConvertedDocument.pdf";

                    //using (var document = DocX.Load(openFileDialog.FileName))
                    //{
                    //    DocX.ConvertToPdf(document, pdfName);
                    //} // не работает так, как надо, но может быть, проблема с границами
                    // нужно попробовать на другом документе !

                    splitPDF(path, pdfName, path);
                }
            }
        }

        private void wordTable(string[,] value, string configPath, int Rows, int Columns)
        {
            int row, col, indOfCom;
            //начало работы с конфигом
            string textConfig = System.String.Empty;
            StreamReader sr = new StreamReader(configPath);
            //создание docx документа
            //выбор папки для создания дока
            //int confPathLength = configPath.Length;
            string docPath = configPath.Remove(configPath.Length - 1, 10) + "wdoc.docx";
            var document = DocX.Create(docPath);
            var table = document.AddTable(Rows, Columns);
            //int r = 0;
            int c = 0;
            for (int r = 0; r <= Rows; r++)
            {
                while (!sr.EndOfStream)
                {
                    textConfig = sr.ReadLine();
                    indOfCom = textConfig.IndexOf(',');
                    row = Convert.ToInt32(textConfig.Substring(0, indOfCom));
                    col = Convert.ToInt32(textConfig.Substring(indOfCom + 1)) - 1;
                    table.Rows[row].Cells[col].Paragraphs[0].Append(value[r, c]);
                    c++;
                }
                c = 0;
            }
            document.Save();
        }


        private void ex2w_Click(object sender, EventArgs e)
        {
            var configFile = String.Empty; // путь до конфигурационного файла
            // чтение конфигурационного файла 
            MessageBox.Show("work");
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "conf files (*.txt)|*.txt|All files (*.*)|*.*";
                ofd.FilterIndex = 1;
                ofd.RestoreDirectory = true;
                if (ofd.ShowDialog() != DialogResult.OK) return;
                configFile = ofd.FileName; // записываем путь до конфигурационного файла
            }
            MessageBox.Show("work1");

            string[,] values = null; // двумерный массив значений таблицы excel
            int maxRow = 0;
            int maxCol = 0;
            // читаем экселевский файл, импортируем с него данные
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() != DialogResult.OK) return;
                var fileName = openFileDialog.SafeFileName; // только имя файла без пути
                var path = Path.GetDirectoryName(openFileDialog.FileName);
                Excel.Application xlApp = new Excel.Application(); //Excel
                Excel.Workbook xlWB; //рабочая книга              
                Excel.Worksheet xlSht; //лист Excel   
                xlWB = xlApp.Workbooks.Open(openFileDialog.FileName,
                    Type.Missing, true, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
                xlSht = (Excel.Worksheet)xlWB.Worksheets[1]; // читаем первый лист
                Excel.Range usedRange = xlSht.UsedRange;
                // возможно потребуются для дальнейшего парсинга
                maxRow = usedRange.Rows.Count;
                maxCol = usedRange.Columns.Count;
                values = usedRange.Value2; // здесь все значения
                xlWB.Close(false);
                xlApp.Quit();
                //  MessageBox.Show(values[1, 1].ToString());
                //   MessageBox.Show(values[2, 2].ToString());
                int Rows, Columns; // сделать ввод пользователем количества столбцов и строк
                //temporary
                Rows = 15;
                Columns = 15;
                //
                wordTable(values, configFile, Rows, Columns);
                //MessageBox.Show(configFile);
            }
        }
    }
}
