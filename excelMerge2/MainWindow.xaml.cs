using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
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
    public class ItemData : TextBlock
    {
        string key;
        public ItemData(string InKey) { key = InKey; }
        public string GetKey() { return key; }
    }

    public static class RowUtils
    {
        static public Dictionary<string, IXLRow> RowsToDict(IEnumerable<IXLRow> Rows)
        {
            return Rows.GroupBy(r => r.Cell(1).GetString()) //第一个格的元素
                       .Where(g => (!string.IsNullOrEmpty(g.Key))) //上面那个函数返回的是g.Key
                       .ToDictionary(g => g.Key, g => g.First()); //主键为g.Key的第一行是g.First
        }

        static public void CopyRow(IXLRow sourceRow, IXLRow targetRow)
        {
            int sourceCount = sourceRow.CellsUsed().Count();
            for (int i = 1; i <= sourceCount; i++)
            {
                string sourceValue = sourceRow.Cell(i).GetString();
                targetRow.Cell(i).Value = sourceValue;
            }
        }
    }

    public class SafeRow
    {
        public IXLRow Data;
        public SafeRow(IXLRow InData) { Data = InData; }
        public SafeRow(Dictionary<string, IXLRow> Dict, string Key) { Dict.TryGetValue(Key, out Data); }

        public string GetValue(int i)
        {
            if (Data != null)
            {
                var Cell = Data.Cell(i);
                try
                {
                    return Cell.GetString();
                }
                catch (NotImplementedException)
                {
                    return Cell.CachedValue.ToString();
                }
            }
            else
            {
                return "";
            }
        }

        public int GetCount()
        {
            if(Data != null)
            {
                return Data.CellsUsed().Count();
            }
            else
            {
                return 0;
            }
        }
    }

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        string LeftPath, RightPath;
        XLWorkbook LeftBook,RightBook;
        IXLWorksheet LeftSheet, RightSheet;
        Dictionary<string, IXLRow> LeftRowDict, RightRowDict;
        Dictionary<string, ItemData> LeftItemDict, RightItemDict;
        int SheetId = 1;

        ScrollViewer svLeft, svRight;
        bool bSyncingSv = false;

        ScrollViewer FindScrollViewer(DependencyObject d)
        {
            if (d == null) return null;
            if (d is ScrollViewer) return (ScrollViewer)d;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(d); i++)
            {
                var child = VisualTreeHelper.GetChild(d, i);
                var sv = FindScrollViewer(child);
                if (sv != null)
                {
                    return sv;
                }
            }
            return null;
        }

        void InitScroller()
        {
            svLeft = FindScrollViewer(ListLeft);
            svRight = FindScrollViewer(ListRight);
            if (svLeft != null)
            {
                svLeft.ScrollChanged += ScrollChanged;
            }
            if (svRight != null)
            {
                svRight.ScrollChanged += ScrollChanged;
            }
        }

        void InitData()
        {
            ListLeft.Items.Clear();
            ListRight.Items.Clear();
            if(LeftRowDict != null)
            {
                LeftRowDict.Clear();
                RightRowDict.Clear();
            }
            LeftItemDict = new Dictionary<string, ItemData>();
            RightItemDict = new Dictionary<string, ItemData>();
        }

        void ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (!bSyncingSv)
            {
                var source = (ScrollViewer)sender;
                var target = source == svLeft ? svRight : svLeft;
                try
                {
                    bSyncingSv = true;
                    target.ScrollToVerticalOffset(e.VerticalOffset);
                    target.ScrollToHorizontalOffset(e.HorizontalOffset);
                }
                finally
                {
                    bSyncingSv = false;
                }
            }
        }

        void UpdateList()
        {
            //清空之前数据
            InitScroller();
            InitData();
            //读表
            LeftSheet = LeftBook.Worksheet(SheetId);
            RightSheet = RightBook.Worksheet(SheetId);

            IEnumerable<IXLRow> LeftRows = LeftSheet.RowsUsed().Where(r => !r.IsEmpty());
            IEnumerable<IXLRow> RightRows = RightSheet.RowsUsed().Where(r => !r.IsEmpty());

            LeftRowDict = RowUtils.RowsToDict(LeftRows);
            RightRowDict = RowUtils.RowsToDict(RightRows);

            IEnumerable<string> allKeys = LeftRowDict.Keys.Union(RightRowDict.Keys);
            foreach (string key in allKeys)
            {
                SafeRow lRow = new SafeRow(LeftRowDict, key);
                SafeRow rRow = new SafeRow(RightRowDict, key);
                int lCount = lRow.GetCount();
                int rCount = rRow.GetCount();
                int Count = Math.Max(lCount, rCount);
                //遍历行
                ItemData LeftItem = new ItemData(key);
                ItemData RightItem = new ItemData(key);
                LeftItemDict[key] = LeftItem;
                RightItemDict[key] = RightItem;
                for (int i = 1; i <= Count; i++)
                {
                    string lValue = lRow.GetValue(i);
                    string rValue = rRow.GetValue(i);
                    var lRun = new Run("|" + lValue + "\t");
                    var rRun = new Run("|" + rValue + "\t");
                    if (!string.Equals(lValue, rValue, StringComparison.Ordinal))
                    {
                        //不一样
                        lRun.Foreground = Brushes.Red;
                        rRun.Foreground = Brushes.Red;
                        LeftItem.Inlines.Add(lRun);
                        RightItem.Inlines.Add(rRun);
                    }
                    else
                    {
                        //一样
                        lRun.Foreground = Brushes.Black;
                        rRun.Foreground = Brushes.Black;
                        LeftItem.Inlines.Add(lRun);
                        RightItem.Inlines.Add(rRun);
                    }
                }
                ListLeft.Items.Add(LeftItem);
                ListRight.Items.Add(RightItem);
            }
        }

        private void Show_Click(object sender, RoutedEventArgs e)
        {
            LeftPath = TextLeft.Text;
            RightPath = TextRight.Text;
            LeftBook = new XLWorkbook(LeftPath);
            RightBook = new XLWorkbook(RightPath);
            UpdateList();
        }

        //select
        ItemData GetItemFromDict(ItemData sourceItem, Dictionary<string, ItemData> ItemDict)
        {
            return ItemDict[sourceItem.GetKey()];
        }

        IXLRow GetRowFromDict(ItemData sourceItem, Dictionary<string, IXLRow> RowDict)
        {
            return RowDict[sourceItem.GetKey()];
        }

        private void List_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!bSyncingSv)
            {
                var sourceList = (ListBox)sender;
                bool bLeft = sourceList == ListLeft;
                ListBox targetList = bLeft ? ListRight : ListLeft;
                Dictionary<string, ItemData> targetItemDict = bLeft ? RightItemDict : LeftItemDict;
                try
                {
                    bSyncingSv = true;
                    if (e.AddedItems.Count > 0)
                    {
                        //添加
                        var sourceItem = (ItemData)e.AddedItems[0];
                        var targetItem = GetItemFromDict(sourceItem, targetItemDict);
                        targetList.SelectedItems.Add(targetItem);
                    }
                    else
                    {
                        //删除
                        var sourceItem = (ItemData)e.RemovedItems[0];
                        var targetItem = GetItemFromDict(sourceItem, targetItemDict);
                        targetList.SelectedItems.Remove(targetItem);
                    }
                }
                finally
                {
                    bSyncingSv = false;
                }
            }
        }

        //merge
        void SyncData(System.Collections.IList sourceItems, bool bLeft)
        {
            Dictionary<string, IXLRow> sourceRowDict = bLeft ? LeftRowDict : RightRowDict;
            Dictionary<string, IXLRow> targetRowDict = bLeft ? RightRowDict : LeftRowDict;
            IXLWorksheet TargetSheet = bLeft ? RightSheet : LeftSheet;
            //把source的ItemDatas转成Rows
            foreach (ItemData i in sourceItems)
            {
                var key = i.GetKey();
                IXLRow sourceRow = sourceRowDict[key]; //source的ItemData转成Row
                IXLRow targetRow;
                targetRowDict.TryGetValue(key, out targetRow);
                int Sub = sourceRow.RowNumber();
                if (targetRow == null)
                {
                    TargetSheet.Row(Sub).InsertRowsAbove(1);
                    targetRow = TargetSheet.Row(Sub);
                }
                RowUtils.CopyRow(sourceRow, targetRow);
            }
            UpdateList();
        }

        private void List_Sync(ListBox sender)
        {
            bool bLeft = sender == ListLeft;
            var sourceItems = sender.SelectedItems;
            SyncData(sourceItems, bLeft);
        }

        private void LeftToRight(object sender, RoutedEventArgs e)
        {
            List_Sync(ListLeft);
        }

        private void RightToLeft(object sender, RoutedEventArgs e)
        {
            List_Sync(ListRight);
        }

        //drop
        private void FilePathTextBox_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void FilePathTextBox_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var paths = ((string[])e.Data.GetData(DataFormats.FileDrop))
                                .Where(p => File.Exists(p) && p.Contains("xls")) //只要表格文件
                                .ToArray();
                if(paths.Length > 0)
                {
                    var tb = (TextBox)sender;
                    tb.Text = paths[0];
                }
                e.Handled = true;
            }
        }

        void CancelReadOnlyAndSave(XLWorkbook Book, string Path)
        {
            var attrs = File.GetAttributes(Path);
            if ((attrs & FileAttributes.ReadOnly) != 0)
            {
                // 清除只读位
                File.SetAttributes(Path, attrs & ~FileAttributes.ReadOnly);
            }
            Book.Save();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            CancelReadOnlyAndSave(LeftBook, LeftPath);
            CancelReadOnlyAndSave(RightBook, RightPath);
        }
    }
}
