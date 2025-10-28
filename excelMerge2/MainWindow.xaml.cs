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

    public class HelperRows
    {
        IEnumerable<IXLRow> Rows;
        List<int> PrimaryKeySubs = new List<int>();
        public HelperRows(IEnumerable<IXLRow> InRows) 
        { 
            Rows = InRows;
            //PrimaryKeySubs.Add(1);
            SetPrimaryKeySubs();
        }

        public Dictionary<string, IXLRow> RowsToDict()
        {
            return Rows.GroupBy(r => GetPrimaryKey(r)) //第一个格的元素
                       .Where(g => (!string.IsNullOrEmpty(g.Key))) //上面那个函数返回的是g.Key
                       .ToDictionary(g => g.Key, g => g.First()); //主键为g.Key的第一行是g.First
        }

        static public void CopyRow(IXLRow sourceRow, IXLRow targetRow)
        {
            int sourceCount = SafeRow.GetCount(sourceRow);
            for (int i = 1; i <= sourceCount; i++)
            {
                string sourceValue = sourceRow.Cell(i).GetString();
                targetRow.Cell(i).Value = sourceValue;
            }
        }
        
        private void SetPrimaryKeySubs()
        {
            PrimaryKeySubs.Clear();
            IXLCells KeyConfigRow = Rows.ElementAt(3).CellsUsed();
            for (int i = 0; i < KeyConfigRow.Count(); i++)
            {
                string KeyConfig = KeyConfigRow.ElementAt(i).GetString();
                if (KeyConfig == "PrimaryKey")
                {
                    PrimaryKeySubs.Add(i + 1); //取的时候是通过IXLRow.Cell(i)取的，从1开始
                }
            }
        }

        public string GetPrimaryKey(IXLRow r)
        {
            string ret = "";
            foreach(int i in PrimaryKeySubs)
            {
                ret += r.Cell(i).GetString();
            }
            return ret;
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
                    //return Cell.CachedValue.ToString();
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

        public static int GetCount(IXLRow InData)
        {
            if (InData != null)
            {
                return InData.LastCellUsed().Address.ColumnNumber;
            }
            else
            {
                return 0;
            }
        }

        public int GetCount()
        {
            return GetCount(Data);
        }
    }

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            //崩溃时输出原因
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                MessageBox.Show(e.ExceptionObject.ToString(), "Fatal");
                Environment.Exit(1);
            };
            InitializeComponent();
            //命令行参数
            string[] CommandArgs = Environment.GetCommandLineArgs();
            if (CommandArgs.Length == 3)
            {
                TextLeft.Text = CommandArgs[1];
                TextRight.Text = CommandArgs[2];
                Show_Click(null, null);
            }
            else if(CommandArgs.Length == 5)
            {
                TextLeft.Text = CommandArgs[2];
                TextRight.Text = CommandArgs[3];
                SaveAsPath = CommandArgs[4];
                Show_Click(null, null);
            }
        }

        string LeftPath, RightPath;
        XLWorkbook LeftBook,RightBook;
        IXLWorksheet LeftSheet, RightSheet;
        Dictionary<string, IXLRow> LeftRowDict, RightRowDict;
        Dictionary<string, ItemData> LeftItemDict, RightItemDict;
        int LeftSheetId = 1, RightSheetId = 1;
        string SaveAsPath = null;

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
            else
            {
                LeftItemDict = new Dictionary<string, ItemData>();
                RightItemDict = new Dictionary<string, ItemData>();
            }
        }

        void InitSheetsList()
        {
            bool bIntersectionIsLeft = LeftBook.Worksheets.Count < RightBook.Worksheets.Count;
            XLWorkbook IntersectionBook = bIntersectionIsLeft ? LeftBook : RightBook;
            XLWorkbook OtherBook = bIntersectionIsLeft ? RightBook : LeftBook;
            //如果要改成union需要处理merge时逻辑：sheet找不到时建表
            int IntersectionSheetId = 1;
            foreach (IXLWorksheet sheet in IntersectionBook.Worksheets)
            {
                TextBlock tb = new TextBlock();
                var r = new Run(sheet.Name);
                //判断两个sheet是否相同
                int OtherSheetId = GetSheetSub(OtherBook, sheet.Name);
                int LeftSheetId, RightSheetId;
                if (bIntersectionIsLeft)
                {
                    LeftSheetId = IntersectionSheetId;
                    RightSheetId = OtherSheetId;
                }
                else
                {
                    LeftSheetId = OtherSheetId;
                    RightSheetId = IntersectionSheetId;
                }
                bool bDiff = IsDiffSheet(LeftSheetId, RightSheetId); //看是否相同
                //设置item UI
                r.Foreground = bDiff ? Brushes.Red : Brushes.Black;
                tb.Inlines.Add(r);
                ListSheet.Items.Add(tb);
                IntersectionSheetId++;
            }
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

        void UpdateList(bool bFirst = false)
        {
            //清空之前数据
            InitScroller();
            InitData();
            if(bFirst)
            {
                ListSheet.Items.Clear();
                LeftSheetId = 1;
                RightSheetId = 1;
                InitSheetsList(); //最右边的sheet列表
            }
            LoadAndUpdateListBySheet(LeftSheetId, RightSheetId); //读表
        }

        bool IsDiffSheet(int InLeftSheetId, int InRightSheetId)
        { 
            return LoadAndUpdateListBySheet(InLeftSheetId, InRightSheetId, true);
        }

        //根据sheet内容刷新UI
        //bReturnWhenFound: true时只返回有无差异，不做表现
        bool LoadAndUpdateListBySheet(int InLeftSheetId, int InRightSheetId, bool bReturnWhenFound = false)
        {
            LeftSheet = LeftBook.Worksheet(InLeftSheetId);
            RightSheet = RightBook.Worksheet(InRightSheetId);

            IEnumerable<IXLRow> LeftRows = LeftSheet.RowsUsed().Where(r => !r.IsEmpty());
            IEnumerable<IXLRow> RightRows = RightSheet.RowsUsed().Where(r => !r.IsEmpty());

            HelperRows LeftHelper = new HelperRows(LeftRows);
            HelperRows RightHelper = new HelperRows(RightRows);
            LeftRowDict = LeftHelper.RowsToDict();
            RightRowDict = RightHelper.RowsToDict();

            return UpdateListBySheet(bReturnWhenFound);
        }

        bool UpdateListBySheet(bool bReturnWhenFound = false)
        {
            bool bOnlyShowDiff = IsShowOnlyDiff();
            IEnumerable<string> allKeys = LeftRowDict.Keys.Union(RightRowDict.Keys);
            int DiffNum = 0;
            foreach (string key in allKeys)
            {
                SafeRow lRow = new SafeRow(LeftRowDict, key);
                SafeRow rRow = new SafeRow(RightRowDict, key);
                int lCount = lRow.GetCount();
                int rCount = rRow.GetCount();
                int Count = Math.Max(lCount, rCount);
                //遍历该行所有字段
                bool bRowHasDiff = false;
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

                        if (bReturnWhenFound)
                        {
                            return true; //找到了不一样的直接返回不一样
                        }
                        else
                        {
                            DiffNum++;
                            bRowHasDiff = true;
                        }
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

                if (!bReturnWhenFound)
                {
                    if (bRowHasDiff || (!bRowHasDiff && !bOnlyShowDiff))
                    {
                        ListLeft.Items.Add(LeftItem);
                        ListRight.Items.Add(RightItem);
                    }
                }
            }

            LabelDiffNum.Content = DiffNum;
            return DiffNum > 0;
        }

        private void Show_Click(object sender, RoutedEventArgs e)
        {
            LeftPath = TextLeft.Text;
            RightPath = TextRight.Text;
            LeftBook = new XLWorkbook(LeftPath);
            RightBook = new XLWorkbook(RightPath);
            UpdateList(true);
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
                        ItemData sourceItem = (ItemData)e.AddedItems[0];
                        ItemData targetItem = GetItemFromDict(sourceItem, targetItemDict);
                        targetList.SelectedItems.Add(targetItem);
                    }
                    else
                    {
                        //删除
                        ItemData sourceItem = (ItemData)e.RemovedItems[0];
                        ItemData targetItem = GetItemFromDict(sourceItem, targetItemDict);
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
                    //之前没有这个主键，需要新增行
                    int targetAppendSub = TargetSheet.RowsUsed().Count() + 1;
                    if (Sub > targetAppendSub) //如果下标很大就加在target末尾
                    {
                        Sub = targetAppendSub;
                    }
                    TargetSheet.Row(Sub).InsertRowsAbove(1);
                    targetRow = TargetSheet.Row(Sub);
                }
                HelperRows.CopyRow(sourceRow, targetRow);
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

        //save
        void CancelReadOnlyAndSave(XLWorkbook Book, string Path)
        {
            var attrs = File.GetAttributes(Path);
            if ((attrs & FileAttributes.ReadOnly) != 0)
            {
                File.SetAttributes(Path, attrs & ~FileAttributes.ReadOnly); //清除只读位
            }
            Book.Save();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (SaveAsPath == null)
            {
                //CancelReadOnlyAndSave(LeftBook, LeftPath);
                CancelReadOnlyAndSave(RightBook, RightPath);
            }
            else
            {
                RightBook.SaveAs(SaveAsPath);
            }
        }

        //多sheet选择
        static public int GetSheetSub(XLWorkbook book, string name)
        {
            int i = 1;
            foreach (IXLWorksheet sheet in book.Worksheets)
            {
                if(sheet.Name == name)
                {
                    return i;
                }
                i++;
            }
            return -1;
        }

        private void Sheet_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var SheetList = (ListBox)sender;
            TextBlock tb = (TextBlock)SheetList.SelectedItem;
            if(tb != null)
            {
                var r = (Run)tb.Inlines.FirstInline;
                string Text = r.Text;
                int NewLeftSheetId = GetSheetSub(LeftBook, Text);
                int NewRightSheetId = GetSheetSub(RightBook, Text);
                if (LeftSheetId != NewLeftSheetId || RightSheetId != NewRightSheetId)
                {
                    LeftSheetId = NewLeftSheetId;
                    RightSheetId = NewRightSheetId;
                    UpdateList();
                }
            }
        }

        //checkBox
        bool IsShowOnlyDiff()
        {
            bool? CheckValue = CheckBoxOnlyDiff.IsChecked;
            if(CheckValue != null)
            {
                return (bool)CheckValue;
            }
            else
            {
                return true;
            }
        }

        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            //清空UI
            ListLeft.Items.Clear();
            ListRight.Items.Clear();
            UpdateListBySheet(false);
        }
    }
}
