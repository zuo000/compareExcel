using System;
using System.IO;
using System.Collections.Generic;
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
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private IWorkbook workbook1_input = null;
        private IWorkbook workbook2_input = null;
        private ISheet sheet_workbook1 = null;
        private ISheet sheet_workbook2 = null;
        private int col_index1 = -1;
        private int col_index1_1 = -1;
        private int col_index2 = -1;
        private int col_index2_1 = -1;
        private bool is_func00_on = false;
        string filename1 = null;
        string filename2 = null;

        public MainWindow()
        {
            InitializeComponent();
        }

        public struct StrData
        {
            public double col2;
            public int row;
        }

        int Rfind_2nd_index(string substr, string str)
        {
            int index = str.LastIndexOf(substr);
            if (index >= 0)
            {
                return str.LastIndexOf(substr, index);
            }
            else
            {
                return -1;
            }
        }

        string Convert_str15_1(string str)
        {
            int index = str.LastIndexOf("00");
            if (index >= 0)
            {
                return str.Remove(index, 2);
            }
            else
            {
                return str;
            }
        }

        string Convert_str15_2(string str)
        {
            int index = Rfind_2nd_index("00", str);
            if (index >= 0)
            {
                return str.Remove(index, 2);
            }
            else
            {
                return str;
            }
        }

        string Convert_str(string str)
        {
            if (str.Length == 16)
            {
                int index = str.LastIndexOf("000");
                if (index >= 0)
                {
                    return str.Remove(index, 3);
                }
                else
                {
                    return str;
                }
            }
            else if (str.Length == 17)
            {
                int index = str.LastIndexOf("0000");
                if (index >= 0)
                {
                    return str.Remove(index, 4);
                }
                else
                {
                    return str;
                }
            }
            else if (str.Length == 13)
            {
                int index = str.Remove(12, 1).LastIndexOf("0");
                if (index > 0)
                {
                    return str.Insert(index, "0");
                }
                else
                {
                    return str;
                }
            }
            else
                return str;
        }

        bool IsSamePrice(double a, double b)
        {
            return (Math.Abs(a-b) > 0.1)?false:true;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SheetBox1.ItemsSource = null;
            ColBox1.ItemsSource = null;
            ColBox1_1.ItemsSource = null;

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel Worksheets|*.xls;*.xlsx";

            if (dialog.ShowDialog() == true)
            {
                TextBox_file1.Text = dialog.FileName;
                filename1 = System.IO.Path.GetFileName(dialog.FileName);

                string extension = System.IO.Path.GetExtension(dialog.FileName);
                FileStream fs = File.Open(dialog.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                if (extension.Equals(".xls"))
                {
                    workbook1_input = new HSSFWorkbook(fs);
                }
                else
                {
                    workbook1_input = new XSSFWorkbook(fs);
                }
                fs.Close();

                List<String> sheetlist1 = new List<string>();
                for (int i = 0; i < workbook1_input.NumberOfSheets; ++i)
                {
                    sheetlist1.Add(workbook1_input.GetSheetName(i));
                }
                SheetBox1.ItemsSource = sheetlist1;
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            SheetBox2.ItemsSource = null;
            ColBox2.ItemsSource = null;
            ColBox2_1.ItemsSource = null;

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel Worksheets|*.xls;*.xlsx";

            if (dialog.ShowDialog() == true)
            {
                TextBox_file2.Text = dialog.FileName;
                filename2 = System.IO.Path.GetFileName(dialog.FileName);

                string extension = System.IO.Path.GetExtension(dialog.FileName);
                FileStream fs = File.Open(dialog.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                if (extension.Equals(".xls"))
                {
                    workbook2_input = new HSSFWorkbook(fs);
                }
                else
                {
                    workbook2_input = new XSSFWorkbook(fs);
                }
                fs.Close();

                List<String> sheetlist2 = new List<string>();
                for (int i = 0; i < workbook2_input.NumberOfSheets; ++i)
                {
                    sheetlist2.Add(workbook2_input.GetSheetName(i));
                }
                SheetBox2.ItemsSource = sheetlist2;
            }
        }

        private void SheetBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SheetBox1.SelectedIndex >= 0)
            {
                sheet_workbook1 = workbook1_input.GetSheetAt(SheetBox1.SelectedIndex);
                IRow row0 = sheet_workbook1.GetRow(0);
                List<string> collist = new List<string>();

                if (row0 != null)
                {
                    for (int i = 0; i < row0.LastCellNum; ++i)
                    {
                        ICell cell = row0.GetCell(i);
                        if (cell != null)
                        {
                            collist.Add(cell.ToString());
                        }
                        else
                        {
                            collist.Add("");
                        }

                    }
                    ColBox1.ItemsSource = collist;
                    ColBox1_1.ItemsSource = collist;
                }
            }
        }

        private void SheetBox2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SheetBox2.SelectedIndex >= 0)
            {
                sheet_workbook2 = workbook2_input.GetSheetAt(SheetBox2.SelectedIndex);
                IRow row0 = sheet_workbook2.GetRow(0);
                List<string> collist = new List<string>();

                if (row0 != null)
                {
                    for (int i = 0; i < row0.LastCellNum; ++i)
                    {
                        ICell cell = row0.GetCell(i);
                        if (cell != null)
                        {
                            collist.Add(cell.ToString());
                        }
                        else
                        {
                            collist.Add("");
                        }

                    }
                    ColBox2.ItemsSource = collist;
                    ColBox2_1.ItemsSource = collist;
                }
            }
        }

        private void ColBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            col_index1 = ColBox1.SelectedIndex;
        }

        private void ColBox2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            col_index2 = ColBox2.SelectedIndex;
        }

        private void ColBox1_1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            col_index1_1 = ColBox1_1.SelectedIndex;
        }

        private void ColBox2_1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            col_index2_1 = ColBox2_1.SelectedIndex;
        }

        private void Compare_Click(object sender, RoutedEventArgs e)
        {
            if (workbook1_input == null)
            {
                MessageBox.Show("请选择文件");
                return;
            }

            if (workbook2_input == null)
            {
                MessageBox.Show("请选择文件");
                return;
            }

            if (sheet_workbook1 == null)
            {
                MessageBox.Show("请选择Sheet");
                return;
            }

            if (sheet_workbook2 == null)
            {
                MessageBox.Show("请选择Sheet");
                return;
            }

            if (col_index1 < 0)
            {
                MessageBox.Show("请选择列");
                return;
            }

            if (col_index2 < 0)
            {
                MessageBox.Show("请选择列");
                return;
            }


            Dictionary<string, StrData> dic1 = new Dictionary<string, StrData>();
            Dictionary<string, StrData> dic2 = new Dictionary<string, StrData>();
            List<int> list_only1 = new List<int>();
            List<int> list_only2 = new List<int>();
            List<string> list_remove = new List<string>();

            for (int i = 1; i <= sheet_workbook1.LastRowNum; ++i)
            {
                IRow row = sheet_workbook1.GetRow(i);

                if (col_index1_1 >= 0)
                {
                    ICell cell = row.GetCell(col_index1_1);
                    StrData data;

                    if (cell.ToString() != "")
                    {
                        data.col2 = Convert.ToDouble(cell.ToString());
                        data.row = row.RowNum;
                    }
                    else
                    {
                        data.col2 = 0;
                        data.row = row.RowNum;
                    }

                    string key = row.GetCell(col_index1).ToString();
                    if (dic1.ContainsKey(key))
                    {
                        dic1[key] = data;
                    }
                    else
                    {
                        dic1.Add(key, data);
                    }
                }
                else
                {
                    StrData data;
                    
                    data.col2 = 0;
                    data.row = row.RowNum;

                    string key = row.GetCell(col_index1).ToString();
                    if (dic1.ContainsKey(key))
                    {
                        dic1[key] = data;
                    }
                    else
                    {
                        dic1.Add(key, data);
                    }
                }
            }

            for (int i = 1; i <= sheet_workbook2.LastRowNum; ++i)
            {
                IRow row = sheet_workbook2.GetRow(i);

                if (col_index2_1 >= 0)
                {
                    ICell cell = row.GetCell(col_index2_1);
                    StrData data;

                    if (cell.ToString() != "")
                    {
                        data.col2 = Convert.ToDouble(cell.ToString());
                        data.row = row.RowNum;
                    }
                    else
                    {
                        data.col2 = 0;
                        data.row = row.RowNum;
                    }
                    
                    string key = row.GetCell(col_index2).ToString();
                    if (dic2.ContainsKey(key))
                    {
                        dic2[key] = data;
                    }
                    else
                    {
                        dic2.Add(key, data);
                    }
                }
                else
                {
                    StrData data;

                    data.col2 = 0;
                    data.row = row.RowNum;

                    string key = row.GetCell(col_index2).ToString();
                    if (dic2.ContainsKey(key))
                    {
                        dic2[key] = data;
                    }
                    else
                    {
                        dic2.Add(key, data);
                    }
                }
            }

            foreach (string str in dic1.Keys)
            {
                if (dic2.ContainsKey(str))
                {
                    if (IsSamePrice(dic1[str].col2, dic2[str].col2))
                    {
                        list_remove.Add(str);
                    }
                }
                else if (is_func00_on == false)
                {
                    list_only1.Add(dic1[str].row);
                }
                else if (str.Length == 15)
                {
                    bool bFind = false;
                    int jj = 0;
                    if (str == "090302002000002")
                    {
                        jj = 1;
                    }
                    jj++;
                    string str_1 = Convert_str15_1(str);
                    string str_2 = Convert_str15_2(str);

                    if (dic2.ContainsKey(str_1))
                    {
                        if (IsSamePrice(dic1[str].col2, dic2[str_1].col2))
                        {
                            list_remove.Add(str_1);
                            bFind = true;
                        }
                    }
                    if (dic2.ContainsKey(str_2))
                    {
                        if (IsSamePrice(dic1[str].col2, dic2[str_2].col2))
                        {
                            list_remove.Add(str_2);
                            bFind = true;
                        }
                    }

                    if (!bFind)
                    {
                        list_only1.Add(dic1[str].row);
                    }
                }
                else
                {
                    string str_tmp = Convert_str(str);
                    if (dic2.ContainsKey(str_tmp))
                    {
                        if (IsSamePrice(dic1[str].col2, dic2[str_tmp].col2))
                        {
                            list_remove.Add(str_tmp);
                        }
                    }
                    else
                    {
                        list_only1.Add(dic1[str].row);
                    }
                }
            }

            foreach (string str in list_remove)
            {
                dic2.Remove(str);
            }

            foreach (string str in dic2.Keys)
            {
                list_only2.Add(dic2[str].row);
            }

            IWorkbook workbook_output = new HSSFWorkbook();
            ISheet sheet1_output = workbook_output.CreateSheet(filename1);
            ISheet sheet2_output = workbook_output.CreateSheet(filename2);

            IRow row0_1 = sheet_workbook1.GetRow(0);
            if (row0_1 != null)
            {
                IRow row_head = sheet1_output.CreateRow(0);
                for (int i = 0; i < row0_1.LastCellNum; ++i)
                {
                    ICell cell = row_head.CreateCell(i);
                    ICell input_cell = row0_1.GetCell(i);
                    if (input_cell != null)
                    {
                        cell.SetCellValue(input_cell.ToString());
                    }
                    else
                    {
                        cell.SetCellValue("");
                    }
                }
            }

            for (int i = 0; i < list_only1.Count; ++i)
            {
                IRow row_input = sheet_workbook1.GetRow(list_only1[i]);
                IRow row_output = sheet1_output.CreateRow(i + 1);
                for (int j = 0; j < row_input.LastCellNum; ++j)
                {
                    ICell cell = row_output.CreateCell(j);
                    ICell input_cell = row_input.GetCell(j);
                    if (input_cell != null)
                    {
                        cell.SetCellValue(input_cell.ToString());
                    }
                    else
                    {
                        cell.SetCellValue("");
                    }
                }
            }

            IRow row0_2 = sheet_workbook2.GetRow(0);
            if (row0_2 != null)
            {
                IRow row_head = sheet2_output.CreateRow(0);
                for (int i = 0; i < row0_2.LastCellNum; ++i)
                {
                    ICell cell = row_head.CreateCell(i);
                    ICell input_cell = row0_2.GetCell(i);
                    if (input_cell != null)
                    {
                        cell.SetCellValue(input_cell.ToString());
                    }
                    else
                    {
                        cell.SetCellValue("");
                    }
                }
            }

            for (int i = 0; i < list_only2.Count; ++i)
            {
                IRow row_input = sheet_workbook2.GetRow(list_only2[i]);
                IRow row_output = sheet2_output.CreateRow(i + 1);
                for (int j = 0; j < row_input.LastCellNum; ++j)
                {
                    ICell cell = row_output.CreateCell(j);
                    ICell input_cell = row_input.GetCell(j);
                    if (input_cell != null)
                    {
                        cell.SetCellValue(input_cell.ToString());
                    }
                    else
                    {
                        cell.SetCellValue("");
                    }
                }
            }

            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = "result";
            dlg.DefaultExt = ".xls";
            dlg.Filter = "Excel Worksheets|*.xls";

            if (dlg.ShowDialog() == true)
            {
                string filename = dlg.FileName;
                using (FileStream fs = File.OpenWrite(filename))
                {
                    workbook_output.Write(fs);
                    fs.Close();

                    MessageBox.Show("保存成功!!");
                }
            }
        }

        private void Func00_Checked(object sender, RoutedEventArgs e)
        {
            is_func00_on = true;
        }
    }
}
