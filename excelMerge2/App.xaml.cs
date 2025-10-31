using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ClosedXML.Excel;

namespace excelMerge2
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>

    public class SafeRow
    {
        public IXLRow Data;
        public SafeRow(IXLRow InData) { Data = InData; }
        public SafeRow(Dictionary<string, IXLRow> Dict, string Key) { Dict.TryGetValue(Key, out Data); }

        public static string GetValue(IXLCell Cell)
        {
            try
            {
                string cacheStr = Cell.CachedValue.ToString();
                if (cacheStr != "")
                {
                    return cacheStr;
                }
                else
                {
                    return Cell.GetString();
                }
            }
            catch (NotImplementedException)
            {
                return "";
            }
        }

        public string GetValue(int i)
        {
            if (Data != null)
            {
                IXLCell Cell = Data.Cell(i);
                return GetValue(Cell);
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

    public class ScrollSyncer
    {
        ListBox ListLeft, ListRight;
        public ScrollSyncer(ListBox InListLeft, ListBox InListRight)
        {
            ListLeft = InListLeft;
            ListRight = InListRight;
        }

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

        public void InitScroller()
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
    }

    public partial class App : Application
    {
        public IXLWorksheet LeftSheet, RightSheet;

        public static App GetApp()
        {
            return (App)Current;
        }
    }
}
