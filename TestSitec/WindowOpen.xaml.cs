using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using TestSitec.Models;
using Path = System.IO.Path;

namespace TestSitec
{
    /// <summary>
    /// Логика взаимодействия для WindowOpen.xaml
    /// </summary>
    public partial class WindowOpen : Window
    {
        private string textFile = default;

        List<TModel> countRKK = new List<TModel>();
        List<TModel> countAppeals = new List<TModel>();
        internal List<TModel> countSum = new List<TModel>();
        internal Stopwatch stopWatch = new Stopwatch();

        public WindowOpen()
        {
            InitializeComponent();            
        }

        private void OpenRKK_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileRKK = new OpenFileDialog()
            {
                Filter = "Text Document (*.txt) | *.txt"
            };

            if (openFileRKK.ShowDialog() == true)
            {
                var crutchesRKK = StrInlist(openFileRKK);
                FileRKK.Text = Path.GetFileName(openFileRKK.FileName);
                
                foreach (var q in crutchesRKK)
                {
                    countRKK.Add(new TModel() { FIO = q.FIO, RKK = q.CountInt, Appeals = 0});
                }
            }
        }
        
        private void OpenAppeals_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileAppeals = new OpenFileDialog()
            {
                Filter = "Text Document (*.txt) | *.txt"
            };

            if (openFileAppeals.ShowDialog() == true)
            {
                var crutchesApp = StrInlist(openFileAppeals);
                FileAppeals.Text = Path.GetFileName(openFileAppeals.FileName);

                foreach (var q in crutchesApp)
                {
                    countAppeals.Add(new TModel() { FIO = q.FIO, RKK = 0, Appeals = q.CountInt });
                }
            }
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (FileRKK.Text.Length == 0 | FileAppeals.Text.Length == 0)
            {
                MessageBox.Show("Выберите оба файла");
            }
            else
            {
                stopWatch.Start();

                var allCnt = countRKK.Union(countAppeals);
                var sumCnt = from r in allCnt
                             group r by r.FIO
                             into newR
                             orderby newR.Key
                             select new
                             {
                                 key = newR.Key,
                                 cntRkk = newR.Sum(x => x.RKK),
                                 cntOb = newR.Sum(x => x.Appeals),
                                 cntSum = newR.Sum(x => x.RKK + x.Appeals)
                             };

                var sumCntGroup = from r in sumCnt
                                  orderby r.key, r.cntRkk, r.cntOb, r.cntSum
                                  select r;

                int i = 1;
                foreach (var q in sumCnt)
                {
                    countSum.Add(new TModel() { FIO = q.key, RKK = q.cntRkk, Appeals = q.cntOb, Sum = q.cntSum, Count = i++ });
                }

                Close();
            }
        }

        // Преобразование строки в Коллекцию
        private List<Crutch> StrInlist(OpenFileDialog openFile)
        {
            textFile = null;

            using (StreamReader sr = new StreamReader(File.OpenRead(openFile.FileName)))
            {
                while (true)
                {
                    string line = sr.ReadLine();
                    if (line == null) break;
                    textFile += line + "\n";
                }
            }

            string oldValue = "(Отв.)";
            string oldValue2 = "(Отв.);";
            string newValue = " ";
            string key = "Климов Сергей Александрович";

            List<Model> list = new List<Model>();

            using (StringReader sr = new StringReader(textFile))
            {
                while (true)
                {
                    string line = sr.ReadLine();
                    if (line == null) break;
                    string[] subs = line.Split('\t');
                    string[] subsReplace = subs[1].ToString().Split(' ');
                    for (int i = 0; i < subsReplace.Length; i++)
                    {
                        if (subsReplace[i] == oldValue | subsReplace[i] == oldValue2)
                        {
                            subsReplace[i] = newValue;
                        }
                    }

                    string[] subsRemoveRuk = subs[0].ToString().Split(' ');

                    if (subs[0].ToString().Split(' ').Length == 3)
                    {
                        subsRemoveRuk[1] = subsRemoveRuk[1].Remove(1) + ".";
                        subsRemoveRuk[2] = subsRemoveRuk[2].Remove(1) + ".";
                    }

                    if (subs[0] == key)
                    {
                        string insteadOf = subsReplace[0] + " " + subsReplace[1];

                        if (insteadOf.EndsWith(";"))
                        {
                            insteadOf = insteadOf.Remove(insteadOf.Length - 1);
                        }

                        list.Add(new Model() { Ruk = insteadOf, Pod = subs[1] });

                        continue;
                    }

                    list.Add(new Model() { Ruk = subsRemoveRuk[0] + " " + subsRemoveRuk[1] + subsRemoveRuk[2], Pod = subs[1] });
                }
            }

            var anonim = from q in list
                         group q by q.Ruk into newQ
                         orderby newQ.Key
                         select new { key = newQ.Key, cnt = newQ.Count() };

            List<Crutch> crutches = new List<Crutch>();

            foreach (var q in anonim)
            {
                crutches.Add(new Crutch() { FIO = q.key, CountInt = q.cnt });
            }

            return crutches;
        }
    }
}
