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
using System.Windows.Shapes;
using ClosedXML.Excel;

namespace excelMerge2
{
    /// <summary>
    /// colDiff.xaml 的交互逻辑
    /// </summary>
    public class TextItemData : TextBlock
    {
        string key;
        public TextItemData(string InKey, string Content) 
        { 
            key = InKey;
            Text = Content;
        }
        public string GetKey() { return key; }
    }

    public partial class colDiff : Window
    {
        public colDiff()
        {
            InitializeComponent();
            UpdateList();
        }

        //记录ItemData
        Dictionary<string, TextItemData> LeftItemDict = new Dictionary<string, TextItemData>();
        Dictionary<string, TextItemData> RightItemDict = new Dictionary<string, TextItemData>();
        //记录colNumber
        Dictionary<string, IXLColumn> LeftValueToColDict;
        Dictionary<string, IXLColumn> RightValueToColDict;

        ScrollSyncer Scroller; //滚动同步
        bool bSyncingSv = false; //左右同步中标记

        void UpdateList()
        {
            //重置数据
            ListLeft.Items.Clear();
            ListRight.Items.Clear();
            LeftItemDict.Clear();
            RightItemDict.Clear();
            if (LeftValueToColDict != null)
            {
                LeftValueToColDict.Clear();
                RightValueToColDict.Clear();
            }
            //初始化滚动同步
            Scroller = new ScrollSyncer(ListLeft, ListRight);
            Scroller.InitScroller();

            UpdateListRaw();
        }

        //diff
        public static string GetColFirstValue(IXLWorksheet sheet, int i)
        {
            IXLCell cell = sheet.Cell(1, i);
            return SafeRow.GetValue(cell);
        }

        public static Dictionary<string, IXLColumn> GetValueToNumberDict(IXLWorksheet sheet)
        {
            var ret = new Dictionary<string, IXLColumn>();
            int count = sheet.LastColumnUsed().ColumnNumber();
            for (int i = 1; i <= count; i++)
            {
                string value = GetColFirstValue(sheet, i);
                IXLColumn col = sheet.Column(i);
                ret[value] = col;
            }
            return ret;
        }

        void UpdateListRaw()
        {
            int lCount = App.GetApp().LeftSheet.LastColumnUsed().ColumnNumber();
            int rCount = App.GetApp().RightSheet.LastColumnUsed().ColumnNumber();
            bool bSourceLeft = lCount >= rCount;
            //建target的 value->number 映射
            LeftValueToColDict = GetValueToNumberDict(App.GetApp().LeftSheet);
            RightValueToColDict = GetValueToNumberDict(App.GetApp().RightSheet);
            Dictionary<string, IXLColumn>.KeyCollection leftValues = LeftValueToColDict.Keys;
            Dictionary<string, IXLColumn>.KeyCollection rightValues = RightValueToColDict.Keys;
            IEnumerable<string> AllValues = leftValues.Union(rightValues);
            //进行diff
            foreach (string value in AllValues)
            {
                string leftContent = leftValues.Contains(value) ? value : "";
                string rightContent = rightValues.Contains(value) ? value : "";
                TextItemData leftItemData = new TextItemData(value, leftContent);
                TextItemData rightItemData = new TextItemData(value, rightContent);
                LeftItemDict[value] = leftItemData;
                RightItemDict[value] = rightItemData;
                if (leftContent == rightContent)
                {
                    //一样
                    leftItemData.Foreground = Brushes.Black;
                    rightItemData.Foreground = Brushes.Black;
                }
                else
                {
                    //不一样
                    leftItemData.Foreground = Brushes.Red;
                    rightItemData.Foreground = Brushes.Red;
                }
                ListLeft.Items.Add(leftItemData);
                ListRight.Items.Add(rightItemData);
            }
        }

        //sync
        static public void CopyCol(IXLColumn sourceCol, IXLColumn targetCol)
        {
            int sourceCount = sourceCol.LastCellUsed().Address.RowNumber;
            for (int i = 1; i <= sourceCount; i++)
            {
                IXLCell cell = sourceCol.Cell(i);
                string sourceValue = SafeRow.GetValue(cell);
                targetCol.Cell(i).Value = sourceValue;
            }
        }

        void SyncData(System.Collections.IList sourceItems, bool bLeft)
        {
            Dictionary<string, IXLColumn> sourceValueToColDict = bLeft ? LeftValueToColDict : RightValueToColDict;
            Dictionary<string, IXLColumn> targetValueToColDict = bLeft ? RightValueToColDict : LeftValueToColDict;
            IXLWorksheet TargetSheet = bLeft ? App.GetApp().RightSheet : App.GetApp().LeftSheet;

            foreach (TextItemData i in sourceItems)
            {
                string value = i.GetKey();
                IXLColumn sourceCol = sourceValueToColDict[value];
                IXLColumn targetCol;
                targetValueToColDict.TryGetValue(value, out targetCol);
                int Sub = sourceCol.ColumnNumber();
                if (targetCol == null)
                {
                    //之前没有这个一列，需要新增
                    int targetAppendSub = TargetSheet.ColumnsUsed().Count() + 1;
                    if (Sub > targetAppendSub) //如果下标很大就加在target末尾
                    {
                        Sub = targetAppendSub;
                    }
                    TargetSheet.Column(Sub).InsertColumnsBefore(1);
                    targetCol = TargetSheet.Column(Sub);
                }
                CopyCol(sourceCol, targetCol);
            }
            //刷新UI
            ((MainWindow)Owner).UpdateList();
            UpdateList();
        }

        public static TextItemData GetItemFromDict(TextItemData sourceItem, Dictionary<string, TextItemData> ItemDict)
        {
            return ItemDict[sourceItem.GetKey()];
        }

        private void List_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!bSyncingSv)
            {
                var sourceList = (ListBox)sender;
                bool bLeft = sourceList == ListLeft;
                ListBox targetList = bLeft ? ListRight : ListLeft;
                Dictionary<string, TextItemData> targetItemDict = bLeft ? RightItemDict : LeftItemDict;
                bSyncingSv = true;
                if (e.AddedItems.Count > 0)
                {
                    //添加
                    TextItemData sourceItem = (TextItemData)e.AddedItems[0];
                    TextItemData targetItem = GetItemFromDict(sourceItem, targetItemDict);
                    targetList.SelectedItems.Add(targetItem);
                }
                else
                {
                    //删除
                    TextItemData sourceItem = (TextItemData)e.RemovedItems[0];
                    TextItemData targetItem = GetItemFromDict(sourceItem, targetItemDict);
                    targetList.SelectedItems.Remove(targetItem);
                }
                bSyncingSv = false;
            }
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
    }
}
