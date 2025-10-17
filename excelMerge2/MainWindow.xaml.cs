using System;
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

using ClosedXML.Excel;

namespace excelMerge2
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        XLWorkbook LeftBook;
        XLWorkbook RightBook;

        static public Dictionary<string, IXLRow> RowsToDict(IEnumerable<IXLRow> Rows)
        {
            return Rows.GroupBy(r => r.Cell(1).GetString()) //第一个格的元素
                       .Where(g => (!string.IsNullOrEmpty(g.Key))) //上面那个函数返回的是g.Key
                       .ToDictionary(g => g.Key, g => g.First()); //主键为g.Key的第一行是g.First
        }

        private void Show_Click(object sender, RoutedEventArgs e)
        {
            //清空之前list
            ListLeft.Items.Clear();
            ListRight.Items.Clear();
            //读表
            LeftBook = new XLWorkbook("C:\\SG_CIV2\\trunk\\CivGame\\Table\\Base\\C创角配置.xlsx");
            RightBook = new XLWorkbook("C:\\SG_CIV2\\branches\\CBT2\\Table\\Base\\C创角配置.xlsx");
            IXLWorksheet LeftSheet = LeftBook.Worksheet(1);
            IXLWorksheet RightSheet = RightBook.Worksheet(1);

            IEnumerable<IXLRow> LeftRows = LeftSheet.RowsUsed().Where(r => !r.IsEmpty());
            IEnumerable<IXLRow> RightRows = RightSheet.RowsUsed().Where(r => !r.IsEmpty());

            Dictionary<string, IXLRow> LeftDict = RowsToDict(LeftRows);
            Dictionary<string, IXLRow> RightDict = RowsToDict(RightRows);

            //缺的整行
            Dictionary<string, IXLRow> LeftDiffRight = LeftDict.Keys.Except(RightDict.Keys).ToDictionary(k => k, k => LeftDict[k]);
            Dictionary<string, IXLRow> RightDiffLeft = RightDict.Keys.Except(LeftDict.Keys).ToDictionary(k => k, k => RightDict[k]);
            //并集逐个比较
            IEnumerable<string> CommonKeys = LeftDict.Keys.Intersect(RightDict.Keys);
            foreach (string key in CommonKeys) //遍历所有key
            {
                IXLRow lRow = LeftDict[key];
                IXLRow rRow = RightDict[key];
                int lCount = lRow.CellsUsed().Count();
                int rCount = rRow.CellsUsed().Count();
                int Count = Math.Max(lCount, rCount);
                //遍历行
                var LeftText = new TextBlock();
                var RightText = new TextBlock();
                for (int i = 1; i <= Count; i++)
                {
                    string lValue = lRow.Cell(i).GetString();
                    string rValue = rRow.Cell(i).GetString();
                    var LeftRun = new Run("|" + lValue + "\t");
                    var RightRun = new Run("|" + rValue + "\t");
                    if (!string.Equals(lValue, rValue, StringComparison.Ordinal))
                    {
                        //不一样
                        LeftRun.Foreground = Brushes.Red;
                        RightRun.Foreground = Brushes.Red;
                        LeftText.Inlines.Add(LeftRun);
                        RightText.Inlines.Add(RightRun);
                    }
                    else
                    {
                        //一样
                        LeftRun.Foreground = Brushes.Black;
                        RightRun.Foreground = Brushes.Black;
                        LeftText.Inlines.Add(LeftRun);
                        RightText.Inlines.Add(RightRun);
                    }
                }
                ListLeft.Items.Add(LeftText);
                ListRight.Items.Add(RightText);
            }
        }

        private void ListLeft_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var lb = (ListBox)sender;
            var item = lb.SelectedItem;
            var a = 1;
        }

        private void ListRight_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var lb = (ListBox)sender;
        }
    }
}
