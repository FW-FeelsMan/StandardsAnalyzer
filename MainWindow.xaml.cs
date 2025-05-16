using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace StandardsAnalyzer
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<string> file1Standards = new List<string>();
        private List<string> file2Standards = new List<string>();
        private BackgroundWorker worker;
        private int totalCells = 0;
        private int processedCells = 0;
        private bool isFile1Processing = false;

        public MainWindow()
        {
            InitializeComponent();
            InitializeBackgroundWorker();
        }

        private void InitializeBackgroundWorker()
        {
            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.WorkerSupportsCancellation = true;
            worker.DoWork += Worker_DoWork;
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
        }

        private void BtnSelectFile1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
            openFileDialog.Title = "Выберите первый Excel файл";

            if (openFileDialog.ShowDialog() == true)
            {
                txtFile1Path.Text = openFileDialog.FileName;
            }
        }

        private void BtnSelectFile2_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
            openFileDialog.Title = "Выберите второй Excel файл";

            if (openFileDialog.ShowDialog() == true)
            {
                txtFile2Path.Text = openFileDialog.FileName;
            }
        }

        private void BtnAnalyze_Click(object sender, RoutedEventArgs e)
        {
            // Проверка наличия файлов
            if (string.IsNullOrEmpty(txtFile1Path.Text) || string.IsNullOrEmpty(txtFile2Path.Text))
            {
                MessageBox.Show("Пожалуйста, выберите оба Excel файла для анализа.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Очистка предыдущих результатов
            lstFile1Standards.Items.Clear();
            lstFile2Standards.Items.Clear();
            file1Standards.Clear();
            file2Standards.Clear();

            // Сброс прогресса
            progressBar.Value = 0;
            lblProgress.Content = "0%";

            // Отключение кнопок на время анализа
            btnSelectFile1.IsEnabled = false;
            btnSelectFile2.IsEnabled = false;
            btnAnalyze.IsEnabled = false;

            // Запуск асинхронной обработки
            if (!worker.IsBusy)
            {
                worker.RunWorkerAsync(new string[] { txtFile1Path.Text, txtFile2Path.Text });
            }
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            string[] filePaths = e.Argument as string[];
            if (filePaths == null || filePaths.Length != 2)
            {
                e.Result = "Ошибка: неверные параметры";
                return;
            }

            try
            {
                // Обработка первого файла
                isFile1Processing = true;
                worker.ReportProgress(0, "Анализ файла 1...");
                List<string> standards1 = ProcessExcelFileAsync(filePaths[0], worker);

                // Обработка второго файла
                isFile1Processing = false;
                worker.ReportProgress(50, "Анализ файла 2...");
                processedCells = 0; // Сброс счетчика для второго файла
                List<string> standards2 = ProcessExcelFileAsync(filePaths[1], worker);

                e.Result = new List<string>[] { standards1, standards2 };
            }
            catch (Exception ex)
            {
                e.Result = "Ошибка: " + ex.Message;
            }
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Обновление прогресс-бара
            progressBar.Value = e.ProgressPercentage;
            lblProgress.Content = $"{e.ProgressPercentage}%";

            // Если передан стандарт, добавляем его в соответствующий список
            if (e.UserState is string && ((string)e.UserState).StartsWith("STANDARD:"))
            {
                string standard = ((string)e.UserState).Substring(9);
                if (isFile1Processing)
                {
                    if (!file1Standards.Contains(standard))
                    {
                        file1Standards.Add(standard);
                        lstFile1Standards.Items.Add(standard);
                    }
                }
                else
                {
                    if (!file2Standards.Contains(standard))
                    {
                        file2Standards.Add(standard);
                        lstFile2Standards.Items.Add(standard);
                    }
                }
            }
            else if (e.UserState is string)
            {
                // Обновление статуса
                Title = e.UserState as string;
            }
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // Включение кнопок после завершения анализа
            btnSelectFile1.IsEnabled = true;
            btnSelectFile2.IsEnabled = true;
            btnAnalyze.IsEnabled = true;

            if (e.Error != null)
            {
                MessageBox.Show("Произошла ошибка при анализе файлов: " + e.Error.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (e.Result is string)
            {
                MessageBox.Show(e.Result as string, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (e.Result is List<string>[])
            {
                List<string>[] results = e.Result as List<string>[];
                file1Standards = results[0];
                file2Standards = results[1];

                // Сортировка списков
                file1Standards = file1Standards.OrderBy(s => s).ToList();
                file2Standards = file2Standards.OrderBy(s => s).ToList();

                // Обновление списков в UI
                lstFile1Standards.Items.Clear();
                lstFile2Standards.Items.Clear();

                foreach (var standard in file1Standards)
                {
                    lstFile1Standards.Items.Add(standard);
                }

                foreach (var standard in file2Standards)
                {
                    lstFile2Standards.Items.Add(standard);
                }

                // Сохранение отчета
                SaveReport();

                Title = "Анализатор стандартов ГОСТ и DIN";
                MessageBox.Show("Анализ завершен. Отчет сохранен.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private List<string> ProcessExcelFileAsync(string filePath, BackgroundWorker worker)
        {
            HashSet<string> standards = new HashSet<string>();
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                workbook = excelApp.Workbooks.Open(filePath, ReadOnly: true);

                // Подсчет общего количества ячеек для прогресса
                totalCells = 0;
                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    Excel.Range usedRange = worksheet.UsedRange;
                    totalCells += usedRange.Rows.Count * usedRange.Columns.Count;
                }

                processedCells = 0;

                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    Excel.Range usedRange = worksheet.UsedRange;

                    for (int row = 1; row <= usedRange.Rows.Count; row++)
                    {
                        for (int col = 1; col <= usedRange.Columns.Count; col++)
                        {
                            if (worker.CancellationPending)
                            {
                                return standards.ToList();
                            }

                            Excel.Range cell = usedRange.Cells[row, col];
                            if (cell.Value != null)
                            {
                                string cellValue = cell.Value.ToString();
                                List<string> foundStandards = ExtractStandards(cellValue);

                                foreach (var std in foundStandards)
                                {
                                    if (!standards.Contains(std))
                                    {
                                        standards.Add(std);
                                        // Отправляем найденный стандарт в UI
                                        worker.ReportProgress(CalculateProgress(), "STANDARD:" + std);
                                    }
                                }
                            }

                            processedCells++;

                            // Обновляем прогресс каждые 10 ячеек для производительности
                            if (processedCells % 10 == 0)
                            {
                                worker.ReportProgress(CalculateProgress());
                            }
                        }
                    }
                }
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return standards.OrderBy(s => s).ToList();
        }

        private int CalculateProgress()
        {
            if (totalCells == 0)
                return 0;

            int baseProgress = isFile1Processing ? 0 : 50;
            int fileProgress = (int)((double)processedCells / totalCells * 50);
            return baseProgress + fileProgress;
        }

        private List<string> ExtractStandards(string text)
        {
            List<string> result = new List<string>();

            if (string.IsNullOrEmpty(text))
                return result;

            // Поиск ГОСТ
            Regex gostPattern = new Regex(@"ГОСТ\s*[-№]?\s*\d+(?:[-–—.]\d+)?(?:[-–—]\d+)?", RegexOptions.IgnoreCase);
            MatchCollection gostMatches = gostPattern.Matches(text);
            foreach (Match match in gostMatches)
            {
                result.Add(match.Value.ToUpper());
            }

            // Поиск DIN
            Regex dinPattern = new Regex(@"DIN\s*[-№]?\s*\d+(?:[-–—.]\d+)?(?:[-–—]\d+)?", RegexOptions.IgnoreCase);
            MatchCollection dinMatches = dinPattern.Matches(text);
            foreach (Match match in dinMatches)
            {
                result.Add(match.Value.ToUpper());
            }

            return result;
        }

        private void SaveReport()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Text Files|*.txt";
            saveFileDialog.Title = "Сохранить отчет";
            saveFileDialog.FileName = "Отчет_по_стандартам.txt";

            if (saveFileDialog.ShowDialog() == true)
            {
                using (StreamWriter writer = new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8))
                {
                    writer.WriteLine("ОТЧЕТ ПО АНАЛИЗУ СТАНДАРТОВ ГОСТ И DIN В EXCEL-ФАЙЛАХ");
                    writer.WriteLine();

                    // Стандарты из первого файла
                    writer.WriteLine("СТАНДАРТЫ, НАЙДЕННЫЕ В ФАЙЛЕ 1");
                    writer.WriteLine();

                    var file1Gost = file1Standards.Where(s => s.Contains("ГОСТ")).ToList();
                    var file1Din = file1Standards.Where(s => s.Contains("DIN")).ToList();

                    writer.WriteLine($"ГОСТы ({file1Gost.Count}):");
                    foreach (var std in file1Gost)
                    {
                        writer.WriteLine($"- {std}");
                    }
                    writer.WriteLine();

                    writer.WriteLine($"DIN ({file1Din.Count}):");
                    foreach (var std in file1Din)
                    {
                        writer.WriteLine($"- {std}");
                    }
                    writer.WriteLine();

                    // Стандарты из второго файла
                    writer.WriteLine("СТАНДАРТЫ, НАЙДЕННЫЕ В ФАЙЛЕ 2");
                    writer.WriteLine();

                    var file2Gost = file2Standards.Where(s => s.Contains("ГОСТ")).ToList();
                    var file2Din = file2Standards.Where(s => s.Contains("DIN")).ToList();

                    writer.WriteLine($"ГОСТы ({file2Gost.Count}):");
                    foreach (var std in file2Gost)
                    {
                        writer.WriteLine($"- {std}");
                    }
                    writer.WriteLine();

                    writer.WriteLine($"DIN ({file2Din.Count}):");
                    foreach (var std in file2Din)
                    {
                        writer.WriteLine($"- {std}");
                    }
                    writer.WriteLine();

                    // Отсутствующие стандарты
                    writer.WriteLine("СТАНДАРТЫ, ОТСУТСТВУЮЩИЕ В ФАЙЛЕ 2");
                    writer.WriteLine();

                    var missingGost = file1Gost.Except(file2Gost).ToList();
                    var missingDin = file1Din.Except(file2Din).ToList();

                    writer.WriteLine($"ГОСТы ({missingGost.Count}):");
                    foreach (var std in missingGost)
                    {
                        writer.WriteLine($"- {std}");
                    }
                    writer.WriteLine();

                    writer.WriteLine($"DIN ({missingDin.Count}):");
                    foreach (var std in missingDin)
                    {
                        writer.WriteLine($"- {std}");
                    }
                }
            }
        }
    }
}
