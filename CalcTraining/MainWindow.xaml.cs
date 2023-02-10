using Microsoft.Win32;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
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

namespace CalcTraining
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        readonly List<Formula> Formulas = new List<Formula>();
        readonly Random random = new Random();
        int Page_Size = 100;
        int Page_Count = 10;

        public MainWindow()
        {
            InitializeComponent();
        }

        void ClearLog()
        {
            textBox.Text = "";
        }

        void Log(string str)
        {
            textBox.AppendText($"{str}\n");
            textBox.ScrollToEnd();
        }

        /// <summary>
        /// [minValue, maxValue)
        /// </summary>
        /// <param name="maxValue"></param>
        /// <param name="minValue"></param>
        /// <returns></returns>
        int GetRandomInteger(int maxValue, int minValue)
        {
            if (minValue >= maxValue)
            {
                throw new ArgumentOutOfRangeException("minValue CANNOT be greater than or equal to maxValue!");
            }

            int r;
            while (true)
            {
                r = random.Next();
                if (maxValue > 0)
                {
                    r %= maxValue;
                }
                if (r < minValue)
                {
                    continue;
                }
                break;
            }
            return r;
        }

        int[] GetRandomIntegers(int count, int maxValue, int minValue)
        {
            int[] result = new int[count];
            for (int i = 0; i < count; i++)
            {
                result[i] = GetRandomInteger(maxValue, minValue);
            }
            return result;
        }

        void Mix_Add_Sub(string name)
        {
            Log($"-----{name}，生成时间：{DateTime.Now}-----");

            Formulas.Clear();

            int[] arr = GetRandomIntegers(Page_Size * Page_Count, 101, 10);
            List<int> list = new List<int>(arr);
            int idx = 0;
            while (list.Count > 0)
            {
                if (idx % Page_Size == 0)
                {
                    Log($"-----第{idx / Page_Size + 1}页：-----");
                }

                int i1 = random.Next(list.Count);
                int x1 = list[i1];
                list.RemoveAt(i1);

                int x2 = GetRandomInteger(x1, 2);

                Formula f = new Formula();
                Formulas.Add(f);
                f.Index = idx % Page_Size + 1;
                if (random.Next() % 100 < 50)
                {
                    f.Number1 = x1 - x2;
                    f.Number2 = x2;
                    f.Operator1 = "+";
                }
                else
                {
                    f.Number1 = x1;
                    f.Number2 = x2;
                    f.Operator1 = "-";
                }
                Log(f.ToString());

                idx++;
            }

            SaveDialog(name);
        }

        void Mix_Mul_Div(string name)
        {
            Log($"-----{name}，生成时间：{DateTime.Now}-----");

            Formulas.Clear();

            int[] arr = GetRandomIntegers(Page_Size * Page_Count, 10, 2);
            List<int> list = new List<int>(arr);
            int idx = 0;
            while (list.Count > 0)
            {
                if (idx % Page_Size == 0)
                {
                    Log($"-----第{idx / Page_Size + 1}页：-----");
                }

                int i1 = random.Next(list.Count);
                int x1 = list[i1];
                list.RemoveAt(i1);

                int x2 = GetRandomInteger(10, 2);

                Formula f = new Formula();
                Formulas.Add(f);
                f.Index = idx % Page_Size + 1;
                if (random.Next() % 100 < 40)
                {
                    f.Number1 = x1;
                    f.Number2 = x2;
                    f.Operator1 = "×";
                }
                else
                {
                    f.Number1 = x1 * x2;
                    f.Number2 = x1;
                    f.Operator1 = "÷";
                }
                Log(f.ToString());

                idx++;
            }

            SaveDialog(name);
        }

        void Mix_All(string name)
        {
            Log($"-----{name}，生成时间：{DateTime.Now}-----");

            Formulas.Clear();

            int count = Page_Size * Page_Count;
            for (int i = 0; i < count; i++)
            {
                if (i % Page_Size == 0)
                {
                    Log($"-----第{i / Page_Size + 1}页：-----");
                }

                Formula f = new Formula();
                Formulas.Add(f);
                f.Index = i % Page_Size + 1;
                int tmp = random.Next() % 100;
                if (tmp < 20)
                {
                    int x1 = GetRandomInteger(101, 11);
                    int x2 = GetRandomInteger(x1, 2);

                    f.Number1 = x1 - x2;
                    f.Number2 = x2;
                    f.Operator1 = "+";
                }
                else if (tmp < 40)
                {
                    int x1 = GetRandomInteger(101, 11);
                    int x2 = GetRandomInteger(x1, 2);

                    f.Number1 = x1;
                    f.Number2 = x2;
                    f.Operator1 = "-";
                }
                else if (tmp < 60)
                {
                    int x1 = GetRandomInteger(10, 2);
                    int x2 = GetRandomInteger(10, 2);

                    f.Number1 = x1;
                    f.Number2 = x2;
                    f.Operator1 = "×";
                }
                else
                {
                    int x1 = GetRandomInteger(10, 2);
                    int x2 = GetRandomInteger(10, 2);

                    f.Number1 = x1 * x2;
                    f.Number2 = x1;
                    f.Operator1 = "÷";
                }
                Log(f.ToString());
            }

            SaveDialog(name);
        }

        void SaveDialog(string name)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Microsoft Excel 文件(*.xls)|*.xls";
            dialog.FileName = $"{name}_{DateTime.Now:yyyyMMddHHmmss}";
            dialog.DefaultExt = "xls";
            dialog.RestoreDirectory = true;
            if (dialog.ShowDialog(this) == true)
            {
                SaveToExcel(name, dialog.FileName);
            }
        }

        void FillPage(Worksheet sheet, string name, int page, int col)
        {
            int pr = (int)Math.Ceiling((double)Page_Size / col);
            int startRow = (pr + 2) * page + 1;
            int endRow = startRow + 2 + pr;


            int cc = 0;
            if (Formulas.Count > 0)
            {
                cc = Formulas[0].ColumnCount;
            }
            sheet.Range[startRow, 1, startRow, col * cc].Merge();
            sheet.Range[startRow, 1].HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range[startRow, 1].VerticalAlignment = VerticalAlignType.Top;
            sheet.Range[startRow, 1].Style.Font.IsBold = true;
            sheet.Range[startRow, 1].Style.Font.Size = 20;
            sheet.Range[startRow, 1].RowHeight = 30;
            sheet.Range[startRow, 1].Text = name;

            sheet.Range[startRow + 1, 1, startRow + 1, col * cc].Merge();
            sheet.Range[startRow + 1, 1].HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range[startRow + 1, 1].VerticalAlignment = VerticalAlignType.Center;
            sheet.Range[startRow + 1, 1].Style.Font.Size = 12;
            sheet.Range[startRow + 1, 1].RowHeight = 21;
            sheet.Range[startRow + 1, 1].Text = "姓名：________     日期：________     用时：______分钟     成绩________";


            for (int i = 0; i < Page_Size; i++)
            {
                int idx = page * Page_Size + i;
                if (idx >= Formulas.Count)
                {
                    break;
                }

                var row = Formulas[idx].ToArray();

                int x = startRow + 2 + i / col;
                for (int j = 0; j < row.Length; j++)
                {
                    int y = (i % col) * cc + j + 1;
                    sheet.Range[x, y].Text = $"{row[j]}";
                }
            }


            sheet.Range[startRow + 2, 1, endRow, col * cc].Style.Font.Size = 12;
            sheet.Range[startRow + 2, 1, endRow, col * cc].RowHeight = 21;
            for (int i = 0; i < col; i++)
            {
                sheet.Range[startRow + 2, i * cc + 1, endRow, i * cc + 1].ColumnWidth = 5.57;
                sheet.Range[startRow + 2, i * cc + 2, endRow, i * cc + 2].ColumnWidth = 5.14;
                sheet.Range[startRow + 2, i * cc + 3, endRow, i * cc + 3].ColumnWidth = 2.29;
                sheet.Range[startRow + 2, i * cc + 4, endRow, i * cc + 4].ColumnWidth = 5.14;
                sheet.Range[startRow + 2, i * cc + 5, endRow, i * cc + 5].ColumnWidth = 2.29;
                sheet.Range[startRow + 2, i * cc + 6, endRow, i * cc + 6].ColumnWidth = 8.57;

                sheet.Range[startRow + 2, i * cc + 1, endRow, i * cc + 2].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range[startRow + 2, i * cc + 3, endRow, i * cc + 3].Style.HorizontalAlignment = HorizontalAlignType.Center;

                sheet.Range[startRow + 2, i * cc + 6, endRow, i * cc + 6].Style.Font.Underline = FontUnderlineType.Single;
            }
        }

        void SaveToExcel(string name, string path, int col = 3)
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;
            sheet.PageSetup.FooterMarginInch = 0.2;
            sheet.PageSetup.CenterFooter = "&\"Arial\"&10&B&K000000第&P页，总&N页";
            sheet.PageSetup.RightFooter = $"&\"Arial\"&8&B&K999999Powered by kim.wu © {DateTime.Now.Year}.";


            for (int i = 0; i < Page_Count; i++)
            {
                FillPage(sheet, name, i, col);
            }


            workbook.SaveToFile(path, ExcelVersion.Version2010);
            workbook.Dispose();

            Log($"文件保存成功：{path}");
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            if (int.TryParse(textBox1.Text, out int s))
            {
                if (s > 0)
                {
                    Page_Size = s;
                }
                else
                {
                    Log("每页题数不能小于0！");
                }
            }
            if (int.TryParse(textBox2.Text, out int c))
            {
                if (c > 0)
                {
                    Page_Count = c;
                }
                else
                {
                    Log("页数不能小于0！");
                }
            }

            if (radioButton.IsChecked == true)
            {
                Mix_Add_Sub($"{radioButton.Content}");
            }
            else if (radioButton1.IsChecked == true)
            {
                Mix_Mul_Div($"{radioButton1.Content}");
            }
            else if (radioButton2.IsChecked == true)
            {
                Mix_All($"{radioButton2.Content}");
            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            ClearLog();
        }
    }
}
