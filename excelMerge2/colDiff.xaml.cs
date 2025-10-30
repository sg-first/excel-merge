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
            //初始化数据
            LeftItemDict.Clear();
            RightItemDict.Clear();
            //初始化滚动同步
            Scroller = new ScrollSyncer(ListLeft, ListRight);
            Scroller.InitScroller();

            UpdateList();
        }

        Dictionary<string, TextItemData> LeftItemDict = new Dictionary<string, TextItemData>();
        Dictionary<string, TextItemData> RightItemDict = new Dictionary<string, TextItemData>();

        ScrollSyncer Scroller; //滚动同步
        bool bSyncingSv = false; //左右同步中标记

        //diff
        public static bool IsLeft(bool bSourceLeft, bool bSource)
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
            IXLWorksheet sheet = bLeft ? App.GetApp().LeftSheet : App.GetApp().RightSheet;
            return sheet;
        }

        ListBox GetSourceOrTargetList(bool bSourceLeft, bool bSource)
        {
            bool bLeft = IsLeft(bSourceLeft, bSource);
            ListBox sheet = bLeft ? ListLeft : ListRight;
            return sheet;
        }

        public static string GetColFirstValue(IXLWorksheet sheet, int i)
        {
            IXLCell cell = sheet.Cell(1, i);
            return SafeRow.GetValue(cell);
        }

        public static Dictionary<string, int> GetValueToNumberDict(IXLWorksheet sheet)
        {
            var ret = new Dictionary<string, int>();
            int count = sheet.LastColumnUsed().ColumnNumber();
            for (int i = 1; i <= count; i++)
            {
                string value = GetColFirstValue(sheet, i);
                ret[value] = i;
            }
            return ret;
        }

        void UpdateList()
        {
            int lCount = App.GetApp().LeftSheet.LastColumnUsed().ColumnNumber();
            int rCount = App.GetApp().RightSheet.LastColumnUsed().ColumnNumber();
            bool bSourceLeft = lCount >= rCount;
            //建target的 value->number 映射
            Dictionary<string, int> leftValueToNumberDict = GetValueToNumberDict(App.GetApp().LeftSheet);
            Dictionary<string, int> rightValueToNumberDict = GetValueToNumberDict(App.GetApp().RightSheet);
            Dictionary<string, int>.KeyCollection leftValues = leftValueToNumberDict.Keys;
            Dictionary<string, int>.KeyCollection rightValues = rightValueToNumberDict.Keys;
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
