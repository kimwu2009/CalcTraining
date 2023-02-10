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

            int r = -1;
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
            Log($"{name}，生成时间：{DateTime.Now}");

            DataTable dt = new DataTable();
            dt.Columns.Add("编号");
            dt.Columns.Add("左");
            dt.Columns.Add("符号");
            dt.Columns.Add("右");
            dt.Columns.Add("等于");
            dt.Columns.Add("结果");

            int[] arr = GetRandomIntegers(Page_Size, 101, 10);
            List<int> list = new List<int>(arr);
            int idx = 1;
            while (list.Count > 0)
            {
                int i1 = random.Next(list.Count);
                int x1 = list[i1];
                list.RemoveAt(i1);

                int x2 = GetRandomInteger(x1, 2);

                if (random.Next() % 100 < 50)
                {
                    Log($"({idx}) {x1 - x2} + {x2} =");
                    dt.Rows.Add($"({idx})", $"{x1 - x2}", "+", $"{x2}", "=", "        ");
                }
                else
                {
                    Log($"({idx}) {x1} - {x2} =");
                    dt.Rows.Add($"({idx})", $"{x1}", "-", $"{x2}", "=", "        ");
                }

                idx++;
            }

            SaveDialog(name, dt);
        }

        void Mix_Mul_Div(string name)
        {
            Log($"{name}，生成时间：{DateTime.Now}");

            DataTable dt = new DataTable();
            dt.Columns.Add("编号");
            dt.Columns.Add("左");
            dt.Columns.Add("符号");
            dt.Columns.Add("右");
            dt.Columns.Add("等于");
            dt.Columns.Add("结果");

            int[] arr = GetRandomIntegers(Page_Size, 10, 2);
            List<int> list = new List<int>(arr);
            int idx = 1;
            while (list.Count > 0)
            {
                int i1 = random.Next(list.Count);
                int x1 = list[i1];
                list.RemoveAt(i1);

                int x2 = GetRandomInteger(10, 2);

                if (random.Next() % 100 < 40)
                {
                    Log($"({idx}) {x1} × {x2} =");
                    dt.Rows.Add($"({idx})", $"{x1}", "×", $"{x2}", "=", "        ");
                }
                else
                {
                    Log($"({idx}) {x1 * x2} ÷ {x1} =");
                    dt.Rows.Add($"({idx})", $"{x1 * x2}", "÷", $"{x1}", "=", "        ");
                }

                idx++;
            }

            SaveDialog(name, dt);
        }

        void Mix_All(string name)
        {
            Log($"{name}，生成时间：{DateTime.Now}");

            DataTable dt = new DataTable();
            dt.Columns.Add("编号");
            dt.Columns.Add("左");
            dt.Columns.Add("符号");
            dt.Columns.Add("右");
            dt.Columns.Add("等于");
            dt.Columns.Add("结果");

            for (int i = 1; i <= Page_Size; i++)
            {
                int tmp = random.Next() % 100;
                if (tmp < 20)
                {
                    int x1 = GetRandomInteger(101, 11);
                    int x2 = GetRandomInteger(x1, 2);
                    Log($"({i}) {x1 - x2} + {x2} =");
                    dt.Rows.Add($"({i})", $"{x1 - x2}", "+", $"{x2}", "=", "        ");
                }
                else if (tmp < 40)
                {
                    int x1 = GetRandomInteger(101, 11);
                    int x2 = GetRandomInteger(x1, 2);
                    Log($"({i}) {x1} - {x2} =");
                    dt.Rows.Add($"({i})", $"{x1}", "-", $"{x2}", "=", "        ");
                }
                else if (tmp < 60)
                {
                    int x1 = GetRandomInteger(10, 2);
                    int x2 = GetRandomInteger(10, 2);
                    Log($"({i}) {x1} × {x2} =");
                    dt.Rows.Add($"({i})", $"{x1}", "×", $"{x2}", "=", "        ");
                }
                else
                {
                    int x1 = GetRandomInteger(10, 2);
                    int x2 = GetRandomInteger(10, 2);
                    Log($"({i}) {x1 * x2} ÷ {x1} =");
                    dt.Rows.Add($"({i})", $"{x1 * x2}", "÷", $"{x1}", "=", "        ");
                }
            }

            SaveDialog(name, dt);
        }

        void SaveDialog(string name, DataTable dt)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Microsoft Excel 文件(*.xls)|*.xls";
            dialog.FileName = $"{name}_{DateTime.Now:yyyyMMddHHmmss}";
            dialog.DefaultExt = "xls";
            dialog.RestoreDirectory = true;
            if (dialog.ShowDialog(this) == true)
            {
                SaveToExcel(name, dialog.FileName, dt);
            }
        }

        void SaveToExcel(string name, string path, DataTable dt, int col = 3)
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;
            sheet.PageSetup.FooterMarginInch = 0.2;
            sheet.PageSetup.CenterFooter = "&\"Arial\"&10&B&K000000第&P页，总&N页";
            sheet.PageSetup.RightFooter = $"&\"Arial\"&8&B&K000000Powered by kim.wu © {DateTime.Now.Year}.";

            int cc = dt.Columns.Count;

            sheet.Range[1, 1, 1, col * cc].Merge();
            sheet.Range[1, 1].HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range[1, 1].VerticalAlignment = VerticalAlignType.Top;
            sheet.Range[1, 1].Style.Font.IsBold = true;
            sheet.Range[1, 1].Style.Font.Size = 20;
            sheet.Range[1, 1].RowHeight = 30;
            sheet.Range[1, 1].Text = name;

            sheet.Range[2, 1, 2, col * cc].Merge();
            sheet.Range[2, 1].HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range[2, 1].VerticalAlignment = VerticalAlignType.Center;
            sheet.Range[2, 1].Style.Font.Size = 12;
            sheet.Range[2, 1].RowHeight = 21;
            sheet.Range[2, 1].Text = "姓名：________     日期：________     用时：______分钟     成绩________";

            int first = 3;
            int rc = dt.Rows.Count;
            for (int i = 0; i < rc; i++)
            {
                var row = dt.Rows[i];

                int x = i / col + first;
                for (int j = 0; j < row.ItemArray.Length; j++)
                {
                    int y = (i % col) * cc + j + 1;
                    sheet.Range[x, y].Text = $"{row.ItemArray[j]}";
                }
            }

            sheet.Range[first, 1, rc / col + first, col * cc].Style.Font.Size = 12;
            sheet.Range[first, 1, rc / col + first, col * cc].RowHeight = 21;
            for (int i = 0; i < col; i++)
            {
                sheet.Range[first, i * cc + 1, rc / col + first, i * cc + 1].ColumnWidth = 5.57;
                sheet.Range[first, i * cc + 2, rc / col + first, i * cc + 2].ColumnWidth = 5.14;
                sheet.Range[first, i * cc + 3, rc / col + first, i * cc + 3].ColumnWidth = 2.29;
                sheet.Range[first, i * cc + 4, rc / col + first, i * cc + 4].ColumnWidth = 5.14;
                sheet.Range[first, i * cc + 5, rc / col + first, i * cc + 5].ColumnWidth = 2.29;
                sheet.Range[first, i * cc + 6, rc / col + first, i * cc + 6].ColumnWidth = 8.57;

                sheet.Range[first, i * cc + 1, rc / col + first, i * cc + 2].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range[first, i * cc + 3, rc / col + first, i * cc + 3].Style.HorizontalAlignment = HorizontalAlignType.Center;

                sheet.Range[first, i * cc + 6, rc / col + first, i * cc + 6].Style.Font.Underline = FontUnderlineType.Single;
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
    }
}
