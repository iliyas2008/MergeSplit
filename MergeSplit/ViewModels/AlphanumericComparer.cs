using System.Text.RegularExpressions;
using System;
using MergeSplit.Models;
using System.Collections.Generic;
using System.Linq;

namespace MergeSplit.ViewModels
{
    public class AlphanumericComparer : IComparer<FileDetails>
    {
        private readonly int columnIndex;

        public AlphanumericComparer(int columnIndex)
        {
            this.columnIndex = columnIndex;
        }

        public int Compare(FileDetails x, FileDetails y)
        {
            string nameX = x.FileName; // Adjust based on the property you want to sort by
            string nameY = y.FileName; // Adjust based on the property you want to sort by

            return AlphanumericCompare(nameX, nameY);
        }

        // Alphanumeric comparison logic
        private static int AlphanumericCompare(string str1, string str2)
        {
            string[] parts1 = Regex.Split(str1, @"(\d+)");
            string[] parts2 = Regex.Split(str2, @"(\d+)");

            int minPartsLength = Math.Min(parts1.Length, parts2.Length);
            for (int i = 0; i < minPartsLength; i++)
            {
                if (parts1[i].All(char.IsDigit) && parts2[i].All(char.IsDigit))
                {
                    int number1 = int.Parse(parts1[i]);
                    int number2 = int.Parse(parts2[i]);
                    int result = number1.CompareTo(number2);

                    if (result != 0)
                    {
                        return result;
                    }
                }
                else
                {
                    int result = string.Compare(parts1[i], parts2[i], StringComparison.OrdinalIgnoreCase);

                    if (result != 0)
                    {
                        return result;
                    }
                }
            }

            return parts1.Length.CompareTo(parts2.Length);
        }
    }
}