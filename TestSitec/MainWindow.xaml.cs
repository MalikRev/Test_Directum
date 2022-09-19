using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;
using TestSitec.Models;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Reflection;

namespace TestSitec
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        List<TModel> models = new List<TModel>();
        string sortItem = "без сортировки";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_ClickLoad(object sender, RoutedEventArgs e)
        {
            Stopwatch stopWatch = new Stopwatch();

            WindowOpen windowOpen = new WindowOpen();

            windowOpen.ShowDialog();
            models = windowOpen.countSum;
            dgTable.ItemsSource = models;

            stopWatch = windowOpen.stopWatch;
            stopWatch.Stop();

            TimeSpan ts = stopWatch.Elapsed;

            tbExecut.Text = "Выполнения алгоритма заняло: " + ts + "\nВыполнено: " + $"{DateTime.Today.ToString("d")}";
        }

        private void Button_ClickSave(object sender, RoutedEventArgs e)
        {
            if (models.Count <= 0)
            {
                MessageBox.Show("Выберите оба файла");
            }
            else
            {
                int allSum = default;
                foreach (var q in models)
                {
                    allSum += q.Sum;
                }

                int rkkSum = default;
                foreach (var q in models)
                {
                    rkkSum += q.RKK;
                }

                int obSum = default;
                foreach (var q in models)
                {
                    obSum += q.Appeals;
                }

                
                var app = new Microsoft.Office.Interop.Word.Application();
                Document doc = app.Documents.Add();

                
                
                Paragraph paragraph1 = doc.Paragraphs.Add();
                paragraph1.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                Range text1 = paragraph1.Range;
                text1.Font.Size = 14;
                text1.Bold = 1;
                text1.Text = "Справка о неисполненных документах и обращениях граждан\n";
                text1.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                text1.InsertParagraphAfter();
                
                Paragraph paragraph2 = doc.Paragraphs.Add();
                Range text2 = paragraph2.Range;
                text2.Font.Size = 10;
                text2.Text = $"Не исполнено в срок {allSum} документов, из них:" + "\n" + 
                             $"- количество неисполненных входящих документов: {rkkSum};" + "\n" + 
                             $"- количество неисполненных письменных обращений граждан: {obSum}." + "\n" + 
                             $"Сортировка: {sortItem}.";
                text2.InsertParagraphAfter();

                Paragraph tableParagraph = doc.Paragraphs.Add();
                Range tableR = tableParagraph.Range;
                Table table = doc.Tables.Add(tableR, models.Count() + 1, 5);
                table.Borders.InsideLineStyle = table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                table.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                Range cellRange = tableParagraph.Range;
                cellRange.Font.Size = 10;
                cellRange = table.Cell(1, 1).Range;
                cellRange.Columns.SetWidth(29, WdRulerStyle.wdAdjustNone);
                cellRange.Font.Size = 10;
                cellRange.Text = "№ п.п.";
                cellRange = table.Cell(1, 2).Range;
                cellRange.Columns.SetWidth(120, WdRulerStyle.wdAdjustNone);
                cellRange.Font.Size = 10;
                cellRange.Text = "Ответственный исполнитель";
                cellRange = table.Cell(1, 3).Range;
                cellRange.Columns.SetWidth(110, WdRulerStyle.wdAdjustNone);
                cellRange.Font.Size = 10;
                cellRange.Text = "Количество неисполненных входящих документов";
                cellRange = table.Cell(1, 4).Range;
                cellRange.Columns.SetWidth(110, WdRulerStyle.wdAdjustNone);
                cellRange.Font.Size = 10;
                cellRange.Text = "Количество неисполненных письменных обращений граждан";
                cellRange = table.Cell(1, 5).Range;
                cellRange.Columns.SetWidth(110, WdRulerStyle.wdAdjustNone);
                cellRange.Font.Size = 10;
                cellRange.Text = "Общее количество документов и обращений";

                table.Rows[1].Range.Bold = 1;
                table.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                int cntTableWord = 1;
                for (int i = 0; i < models.Count(); i++)
                {
                    var currentCat = models[i];

                    cellRange = table.Cell(i + 2, 1).Range;
                    cellRange.Font.Size = 10;
                    cellRange.Text = cntTableWord.ToString();
                    cntTableWord++;
                    cellRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    

                    cellRange = table.Cell(i + 2, 2).Range;
                    cellRange.Font.Size = 10;
                    cellRange.Text = currentCat.FIO;                    

                    cellRange = table.Cell(i + 2, 3).Range;
                    cellRange.Font.Size = 10;
                    cellRange.Text = currentCat.RKK.ToString();
                    cellRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = table.Cell(i + 2, 4).Range;
                    cellRange.Font.Size = 10;
                    cellRange.Text = currentCat.Appeals.ToString();
                    cellRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = table.Cell(i + 2, 5).Range;
                    cellRange.Font.Size = 10;
                    cellRange.Text = currentCat.Sum.ToString();
                    cellRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }

                Paragraph paragraph3 = doc.Paragraphs.Add();
                Range text3 = paragraph3.Range;
                text3.Font.Size = 10;
                text3.Text = $"\nДата составления справки:      {DateTime.Today:d}";
                
                text3.InsertParagraphAfter();

                string path = System.IO.Directory.GetCurrentDirectory() + "\\Отчет.rtf";

                doc.SaveAs(path);
                MessageBox.Show("Файл 'Отчет.rtf' успешно сохранен в корневую папку 'TestSitec.exe'");
                doc.Close();
            }
        }

        private void FIO_Selected(object sender, RoutedEventArgs e)
        {
            if (FIO.IsSelected == true)
            {
                FIO.Content = "";
                RKK.Content = "Количесту РКК";
                OB.Content = "Количеству обращений";
                SUM.Content = "Общему количеству";

                sortItem = "Фамилии";
            }

            var fio = from r in models
                      orderby r.FIO descending
                      select r;

            List<TModel> buffer = new List<TModel>();

            int i = 1;
            foreach (var q in fio)
            {
                buffer.Add(new TModel() { FIO = q.FIO, RKK = q.RKK, Appeals = q.Appeals, Sum = q.Sum, Count = i++ });
            }
            models = buffer;

            dgTable.ItemsSource = models;
        }

        private void RKK_Selected(object sender, RoutedEventArgs e)
        {
            if (RKK.IsSelected == true)
            {
                FIO.Content = "Фамилии";
                RKK.Content = "";
                OB.Content = "Количеству обращений";
                SUM.Content = "Общему количеству";

                sortItem = "Количесту РКК";
            }

            var rkk = from r in models
                      orderby r.RKK descending, r.Appeals descending
                      select r;

            List<TModel> buffer = new List<TModel>();

            int i = 1;
            foreach (var q in rkk)
            {
                buffer.Add(new TModel() { FIO = q.FIO, RKK = q.RKK, Appeals = q.Appeals, Sum = q.Sum, Count =  i++});
            }
            models = buffer;

            dgTable.ItemsSource = models;
        }

        private void OB_Selected(object sender, RoutedEventArgs e)
        {
            if (OB.IsSelected == true)
            {
                FIO.Content = "Фамилии";
                RKK.Content = "Количесту РКК";
                OB.Content = "";
                SUM.Content = "Общему количеству";

                sortItem = "Количеству обращений";
            }            

            var obr = from r in models
                      orderby r.Appeals descending, r.RKK descending
                      select r;

            List<TModel> buffer = new List<TModel>();

            int i = 1;
            foreach (var q in obr)
            {
                buffer.Add(new TModel() { FIO = q.FIO, RKK = q.RKK, Appeals = q.Appeals, Sum = q.Sum, Count = i++ });
            }
            models = buffer;

            dgTable.ItemsSource = models;
        }

        private void SUM_Selected(object sender, RoutedEventArgs e)
        {
            if (SUM.IsSelected == true)
            {
                FIO.Content = "Фамилии";
                RKK.Content = "Количесту РКК";
                OB.Content = "Количеству обращений";
                SUM.Content = "";

                sortItem = "Общему количеству документов";
            }

            var sum = from r in models
                      orderby r.Sum descending, r.RKK descending
                      select r;

            List<TModel> buffer = new List<TModel>();

            int i = 1;
            foreach (var q in sum)
            {
                buffer.Add(new TModel() { FIO = q.FIO, RKK = q.RKK, Appeals = q.Appeals, Sum = q.Sum, Count = i++ });
            }
            models = buffer;

            dgTable.ItemsSource = models;
        }
    }
}

