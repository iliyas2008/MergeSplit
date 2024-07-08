using System.Text.RegularExpressions;
using System.Windows.Controls;
using System;
using System.Collections;

namespace MergeSplit.ViewModels
{
    public class AlphanumericComparer : IComparer
    {
        private readonly int columnIndex;

        public AlphanumericComparer(int columnIndex)
        {
            this.columnIndex = columnIndex;
        }

        public int Compare(object x, object y)
        {
            var itemX = x as ListViewItem;
            var itemY = y as ListViewItem;

            if (itemX == null || itemY == null)
                return 0;

            /*string nameX = itemX.SubItems[columnIndex].Text;
            string nameY = itemY.SubItems[columnIndex].Text;*/
            
            string nameX = itemX.Content as string;
            string nameY = itemY.Content as string;

            return AlphanumericCompare(nameX, nameY);
        }

        private static int AlphanumericCompare(string str1, string str2)
        {
            string[] parts1 = Regex.Split(str1, @"(\d+)");
            string[] parts2 = Regex.Split(str2, @"(\d+)");

            int minPartsLength = Math.Min(parts1.Length, parts2.Length);
            for (int i = 0; i < minPartsLength; i++)
            {
                if (int.TryParse(parts1[i], out int number1) && int.TryParse(parts2[i], out int number2))
                {
                    int result = number1.CompareTo(number2);
                    if (result != 0)
                        return result;
                }
                else
                {
                    int result = string.Compare(parts1[i], parts2[i], StringComparison.OrdinalIgnoreCase);
                    if (result != 0)
                        return result;
                }
            }

            return parts1.Length.CompareTo(parts2.Length);
        }
    }
}