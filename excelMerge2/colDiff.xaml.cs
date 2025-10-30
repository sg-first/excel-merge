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
            //初始化滚动同步
            Scroller = new ScrollSyncer(ListLeft, ListRight);
            Scroller.InitScroller();

            UpdateList();
        }

        ScrollSyncer Scroller; //滚动同步
        bool bSyncingSv = false; //左右同步中标记

        bool IsLeft(bool bSourceLeft, bool bSource)
        {
            if (bSourceLeft)
            {
                if (bSource)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                if (bSource)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }

        IXLWorksheet GetSourceOrTargetSheet(bool bSourceLeft, bool bSource)
        {
            bool bLeft = IsLeft(bSourceLeft, bSource);
            IXLWorksheet sheet = bLeft ? App.GetApp().RightSheet : App.GetApp().LeftSheet;
            return sheet;
        }

        ListBox GetSourceOrTargetList(bool bSourceLeft, bool bSource)
        {
            bool bLeft = IsLeft(bSourceLeft, bSource);
            ListBox sheet = bLeft ? ListLeft : ListRight;
            return sheet;
        }

        string GetColFirstValue(IXLWorksheet sheet, int i)
        {
            IXLCell cell = sheet.Cell(1, i);
            return SafeRow.GetValue(cell);
        }

        void UpdateList()
        {
            int lCount = App.GetApp().LeftSheet.LastColumnUsed().ColumnNumber();
            int rCount = App.GetApp().RightSheet.LastColumnUsed().ColumnNumber();
            bool bSourceLeft = lCount >= rCount;
            int sourceCount = Math.Max(lCount, rCount);
            int targetCount = Math.Min(lCount, rCount);
            //把target的所有值放到set里，后面要查有无
            IXLWorksheet sourceSheet = GetSourceOrTargetSheet(bSourceLeft, true);
            IXLWorksheet targetSheet = GetSourceOrTargetSheet(bSourceLeft, false);
            var targetValueToNumberDict = new Dictionary<string, int>();
            for (int i = 1; i <= targetCount; i++)
            {
                string value = GetColFirstValue(targetSheet, i);
                targetValueToNumberDict[value] = i;
            }
            //进行diff
            ListBox sourceList = GetSourceOrTargetList(bSourceLeft, true);
            ListBox targetList = GetSourceOrTargetList(bSourceLeft, false);
            int targetI = 1;
            for (int i = 1; i <= sourceCount; i++)
            {
                string value = GetColFirstValue(sourceSheet, i);
                TextItemData sourceItemData = new TextItemData(value, value);
                TextItemData targetItemData;
                if (targetValueToNumberDict.TryGetValue(value, out targetI))
                {
                    //相同元素
                    targetItemData = new TextItemData(value, value);
                    sourceItemData.Foreground = Brushes.Black;
                    targetItemData.Foreground = Brushes.Black;
                }
                else
                {
                    //存在差异
                    targetItemData = new TextItemData(value, "");
                    sourceItemData.Foreground = Brushes.Red;
                    targetItemData.Foreground = Brushes.Red;
                }
                sourceList.Items.Add(sourceItemData);
                targetList.Items.Add(targetItemData);
            }
        }

        void SyncData(System.Collections.IList sourceItems, bool bLeft)
        {
            IXLWorksheet TargetSheet = bLeft ? App.GetApp().RightSheet : App.GetApp().LeftSheet;
        }

        private void List_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!bSyncingSv)
            {
                var sourceList = (ListBox)sender;
                bool bLeft = sourceList == ListLeft;
                ListBox targetList = bLeft ? ListRight : ListLeft;

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
