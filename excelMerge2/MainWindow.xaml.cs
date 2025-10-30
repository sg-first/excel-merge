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
using System.Collections.ObjectModel;
using System.Windows.Markup;
using ClosedXML.Excel;

namespace excelMerge2
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public class ItemData
    {
        string key;
        public ItemData(string InKey) { key = InKey; }
        public string GetKey() { return key; }

        public struct CellData
        {
            public string Content { get; set; }
            public Brush Color { get; set; }
        }

        public ObservableCollection<CellData> AllCell { get; } = new ObservableCollection<CellData>();

        public void AddCell(string InContent, Brush InColor)
        {
            AllCell.Add(new CellData { Content = InContent, Color = InColor });
        }
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
            const int PrimaryKeyRowNumber = 3;
            int a = Rows.Count();
            if (Rows.Count() > PrimaryKeyRowNumber)
            {
                IXLCells KeyConfigRow = Rows.ElementAt(PrimaryKeyRowNumber).CellsUsed();
                for (int i = 0; i < KeyConfigRow.Count(); i++)
                {
                    string KeyConfig = KeyConfigRow.ElementAt(i).GetString();
                    if (KeyConfig == "PrimaryKey")
                    {
                        PrimaryKeySubs.Add(i + 1); //取的时候是通过IXLRow.Cell(i)取的，从1开始
                    }
                }
            }    
        }

        public string GetPrimaryKey(IXLRow r)
        {
            if(PrimaryKeySubs.Count > 0)
            {
                string ret = "";
                foreach (int i in PrimaryKeySubs)
                {
                    ret += r.Cell(i).GetString();
                }
                return ret;
            }
            else
            {
                return r.RowNumber().ToString(); //没有主键，用行号当主键
            }
        }

        public IDictionary<int, int> GetMaxLengthDict()
        {
            var maxLengthDict = new Dictionary<int, int>();
            foreach (IXLRow row in Rows)
            {
                foreach (IXLCell Cell in row.CellsUsed())
                {
                    string content = SafeRow.GetValue(Cell);
                    int len = content.Length;

                    int columnNumber = Cell.Address.ColumnNumber;
                    if (maxLengthDict.TryGetValue(columnNumber, out int existing))
                    {
                        if (len > existing)
                        {
                            maxLengthDict[columnNumber] = len;
                        }
                    }
                    else
                    {
                        maxLengthDict[columnNumber] = len;
                    }
                }
            }
            return maxLengthDict;
        }
    }

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            DataContext = this;
            //崩溃时输出原因
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                MessageBox.Show(e.ExceptionObject.ToString(), "Fatal");
                Environment.Exit(1);
            };
            InitializeComponent();
            Scroller = new ScrollSyncer(ListLeft, ListRight); //初始化滚动同步
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

        //表数据
        string LeftPath, RightPath;
        XLWorkbook LeftBook, RightBook;
        int LeftSheetId = 1, RightSheetId = 1;
        string SaveAsPath = null;
        //sheet缓存数据
        public class SheetData
        {
            public Dictionary<string, IXLRow> LeftRowDict, RightRowDict;
            public IEnumerable<string> AllKeys;
            public IDictionary<int, int> LeftMaxLengthDict, RightMaxLengthDict;
            public void SetRowDict(IEnumerable<IXLRow> LeftRows, IEnumerable<IXLRow> RightRows)
            {
                HelperRows LeftHelper = new HelperRows(LeftRows);
                HelperRows RightHelper = new HelperRows(RightRows);
                LeftRowDict = LeftHelper.RowsToDict();
                RightRowDict = RightHelper.RowsToDict();
                AllKeys = LeftRowDict.Keys.Union(RightRowDict.Keys);
                LeftMaxLengthDict = LeftHelper.GetMaxLengthDict();
                RightMaxLengthDict = RightHelper.GetMaxLengthDict();
            }

            public Dictionary<string, ItemData> LeftItemDict = new Dictionary<string, ItemData>();
            public Dictionary<string, ItemData> RightItemDict = new Dictionary<string, ItemData>();

            public void Clear()
            {
                if (LeftRowDict != null)
                {
                    LeftRowDict.Clear();
                    RightRowDict.Clear();
                    //allKeys
                    LeftMaxLengthDict.Clear();
                    RightMaxLengthDict.Clear();
                }
                else
                {
                    LeftItemDict.Clear();
                    RightItemDict.Clear();
                }
            }
        }
        SheetData SheetCacheData = new SheetData();
        //和GridList UI数据绑定的Data
        public ObservableCollection<ItemData> LeftAllItemData { get; } = new ObservableCollection<ItemData>();
        public ObservableCollection<ItemData> RightAllItemData { get; } = new ObservableCollection<ItemData>();
        ScrollSyncer Scroller; //滚动同步
        bool bSyncingSv = false; //左右同步中标记
        colDiff ColDiffWindow = null; //colDiff用的子窗体

        void InitData()
        {
            ClearGridList();
            SheetCacheData.Clear();
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

        void UpdateList(bool bFirst = false)
        {
            //清空之前数据
            InitData();
            Scroller.InitScroller();
            if (bFirst)
            {
                ListSheet.Items.Clear();
                LeftSheetId = 1;
                RightSheetId = 1;
                InitSheetsList(); //最右边的sheet列表
            }
            LoadAndUpdateGridListBySheet(LeftSheetId, RightSheetId); //读表
        }

        bool IsDiffSheet(int InLeftSheetId, int InRightSheetId)
        { 
            return LoadAndUpdateGridListBySheet(InLeftSheetId, InRightSheetId, true);
        }

        //根据sheet内容刷新UI
        //bReturnWhenFound: true时只返回有无差异，不做表现
        bool LoadAndUpdateGridListBySheet(int InLeftSheetId, int InRightSheetId, bool bReturnWhenFound = false)
        {
            App.GetApp().LeftSheet = LeftBook.Worksheet(InLeftSheetId);
            App.GetApp().RightSheet = RightBook.Worksheet(InRightSheetId);

            IEnumerable<IXLRow> LeftRows = App.GetApp().LeftSheet.RowsUsed().Where(r => !r.IsEmpty());
            IEnumerable<IXLRow> RightRows = App.GetApp().RightSheet.RowsUsed().Where(r => !r.IsEmpty());
            SheetCacheData.SetRowDict(LeftRows, RightRows);

            return UpdateGridListBySheet(bReturnWhenFound);
        }

        bool UpdateGridListBySheet(bool bReturnWhenFound = false)
        {
            bool bOnlyShowDiff = IsShowOnlyDiff();
            int DiffNum = 0;
            foreach (string key in SheetCacheData.AllKeys)
            {
                SafeRow lRow = new SafeRow(SheetCacheData.LeftRowDict, key);
                SafeRow rRow = new SafeRow(SheetCacheData.RightRowDict, key);
                int lCount = lRow.GetCount();
                int rCount = rRow.GetCount();
                int Count = Math.Max(lCount, rCount);
                //遍历该行所有字段
                bool bRowHasDiff = false;
                ItemData LeftItem = new ItemData(key);
                ItemData RightItem = new ItemData(key);
                SheetCacheData.LeftItemDict[key] = LeftItem;
                SheetCacheData.RightItemDict[key] = RightItem;
                for (int ColNumber = 1; ColNumber <= Count; ColNumber++)
                {
                    string lValue = lRow.GetValue(ColNumber);
                    string rValue = rRow.GetValue(ColNumber);
                    if (!string.Equals(lValue, rValue, StringComparison.Ordinal))
                    {
                        //不一样
                        LeftItem.AddCell(lValue, Brushes.Red);
                        RightItem.AddCell(rValue, Brushes.Red);

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
                        LeftItem.AddCell(lValue, Brushes.Black);
                        RightItem.AddCell(rValue, Brushes.Black);
                    }
                }

                if (!bReturnWhenFound)
                {
                    if (bRowHasDiff || (!bRowHasDiff && !bOnlyShowDiff))
                    {
                        //把item添加到UI TODO:感觉可以在LeftItemDict都添加完毕后统一搬一遍
                        LeftAllItemData.Add(LeftItem);
                        RightAllItemData.Add(RightItem);
                        //选中所有差异
                        if (bRowHasDiff)
                        {
                            ListLeft.SelectedItems.Add(LeftItem);
                            ListRight.SelectedItems.Add(RightItem);
                        }
                    }
                }
            }

            if (!bReturnWhenFound)
            {
                SetupGridList(LeftGridView);
                SetupGridList(RightGridView);
                LabelDiffNum.Content = DiffNum; //差异数
            }
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

        private void DiffCol_Click(object sender, RoutedEventArgs e)
        {
            if (App.GetApp().LeftSheet != null)
            {
                if (ColDiffWindow == null)
                {
                    ColDiffWindow = new colDiff();
                    ColDiffWindow.Closed += (s, args) => ColDiffWindow = null;
                    ColDiffWindow.Owner = this;
                }
                ColDiffWindow.Show();
            }
        }

        //GridList UI
        private void SetupGridList(GridView gridView)
        {
            gridView.Columns.Clear();
            // 取当前所有行的最大列数
            var AllItemData = gridView == LeftGridView ? LeftAllItemData : RightAllItemData;
            if (AllItemData.Count > 0)
            {
                int maxCols = AllItemData.Max(r => r.AllCell.Count);
                for (int i = 0; i < maxCols; i++)
                {
                    var col = new GridViewColumn { Header = $"{i + 1}" };
                    col.CellTemplate = CreateCellTemplate(i);
                    gridView.Columns.Add(col);
                }
            }
        }

        private DataTemplate CreateCellTemplate(int index)
        {
            // 注意：这里用双大括号 "{{" 来在插值字符串中输出 "{Binding ...}"
            string xaml =
            $@"<DataTemplate xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'>
                <TextBlock Text='{{Binding AllCell[{index}].Content}}'
                           Foreground='{{Binding AllCell[{index}].Color}}'
                           Padding='4,2'
                           VerticalAlignment='Center'/>
            </DataTemplate>";
            return (DataTemplate)XamlReader.Parse(xaml);
        }

        void ClearGridList()
        {
            LeftAllItemData.Clear();
            RightAllItemData.Clear();
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
                Dictionary<string, ItemData> targetItemDict = bLeft ? SheetCacheData.RightItemDict : SheetCacheData.LeftItemDict;
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
            Dictionary<string, IXLRow> sourceRowDict = bLeft ? SheetCacheData.LeftRowDict : SheetCacheData.RightRowDict;
            Dictionary<string, IXLRow> targetRowDict = bLeft ? SheetCacheData.RightRowDict : SheetCacheData.LeftRowDict;
            IXLWorksheet TargetSheet = bLeft ? App.GetApp().RightSheet : App.GetApp().LeftSheet;
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
            ClearGridList();
            UpdateGridListBySheet(false);
        }
    }
}
