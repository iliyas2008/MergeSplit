using IliyasAddIn.Utilities;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Document = Microsoft.Office.Interop.Word.Document;
using Paragraph = Microsoft.Office.Interop.Word.Paragraph;

namespace MergeSplit
{
    public partial class MergeForm : Form
    {
        private List<Document> mergedDocuments;
        private string password;

        public MergeForm()
        {
            InitializeComponent();
            mergedDocuments = Globals.ThisAddIn.mergedDocList;
        }
        private void MergeForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Close any open Word instances
            CloseWordInstances();

            // Close Word application and release resources
            //CloseWordApplication();
        }

        private void CloseWordInstances()
        {
            Process[] processes = Process.GetProcessesByName("WINWORD");
            foreach (Process process in processes)
            {
                try
                {
                    process.Kill();
                    process.WaitForExit();
                }
                catch (Exception ex)
                {
                    Misc.ShowError($"Error closing Word instance:" + ex.Message);
                }
            }
        }

        private void CloseWordApplication()
        {
            // Close all merged documents
            foreach (Document doc in mergedDocuments)
            {
                doc.Close();
            }

            // Release the Word application
            //if (wordApp != null)
            //{
            //    wordApp.Quit();
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            //    wordApp = null;
            //}
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            lvFiles.Items.Clear();
        }

        private void BtnAddFiles_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Word Documents|*.doc;*.docx"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (string filePath in openFileDialog.FileNames)
                {
                    FileInfo fileInfo = new FileInfo(filePath);
                    if (IsFileExtensionValid(fileInfo.Extension))
                    {
                        ListViewItem item = new ListViewItem(fileInfo.Name);
                        item.SubItems.Add(fileInfo.LastWriteTime.ToString());
                        item.SubItems.Add(fileInfo.FullName);
                        lvFiles.Items.Add(item);
                    }
                }
                SortListViewByFileName(0); // Sort by "File Name" column using alphanumeric sorting
            }
        }

        private void BtnAddFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                string folderPath = folderBrowserDialog.SelectedPath;
                string[] files = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                    .Where(file => !Path.GetFileName(file).StartsWith("~"))
                    .ToArray();

                foreach (string filePath in files)
                {
                    FileInfo fileInfo = new FileInfo(filePath);
                    // Exclude temporary files
                    if ((fileInfo.Attributes & FileAttributes.Temporary) == FileAttributes.Temporary)
                    {
                        continue;
                    }

                    string fileExtension = fileInfo.Extension.ToLower();

                    if (IsFileExtensionValid(fileExtension))
                    {
                        ListViewItem item = new ListViewItem(fileInfo.Name);
                        item.SubItems.Add(fileInfo.LastWriteTime.ToString());
                        item.SubItems.Add(fileInfo.FullName);
                        lvFiles.Items.Add(item);
                    }
                }
                SortListViewByFileName(0); // Sort by "File Name" column using alphanumeric sorting
            }
        }
        private bool IsFileExtensionValid(string fileExtension)
        {
            string[] validExtensions = { ".doc", ".docx" };
            return validExtensions.Contains(fileExtension.ToLower());
        }

        private void BtnMoveUp_Click(object sender, EventArgs e)
        {
            if (lvFiles.SelectedItems.Count > 0)
            {
                ListViewItem selectedItem = lvFiles.SelectedItems[0];
                int selectedIndex = selectedItem.Index;

                if (selectedIndex > 0)
                {
                    lvFiles.Items.Remove(selectedItem);
                    lvFiles.Items.Insert(selectedIndex - 1, selectedItem);
                    selectedItem.Selected = true;
                }
            }
        }

        private void BtnMoveDown_Click(object sender, EventArgs e)
        {
            if (lvFiles.SelectedItems.Count > 0)
            {
                ListViewItem selectedItem = lvFiles.SelectedItems[0];
                int selectedIndex = selectedItem.Index;

                if (selectedIndex < lvFiles.Items.Count - 1)
                {
                    lvFiles.Items.Remove(selectedItem);
                    lvFiles.Items.Insert(selectedIndex + 1, selectedItem);
                    selectedItem.Selected = true;
                }
            }
        }

        private void BtnMoveFirst_Click(object sender, EventArgs e)
        {
            // Check if there are any selected items
            if (lvFiles.SelectedItems.Count > 0)
            {
                // Move each selected item to the first position
                foreach (ListViewItem selectedItem in lvFiles.SelectedItems)
                {
                    int currentIndex = selectedItem.Index;
                    if (currentIndex > 0)
                    {
                        lvFiles.Items.Remove(selectedItem);
                        lvFiles.Items.Insert(0, selectedItem);
                    }
                }
            }
        }

        private void BtnMoveLast_Click(object sender, EventArgs e)
        {
            // Check if there are any selected items
            if (lvFiles.SelectedItems.Count > 0)
            {
                // Move each selected item to the last position
                foreach (ListViewItem selectedItem in lvFiles.SelectedItems)
                {
                    int currentIndex = selectedItem.Index;
                    int lastIndex = lvFiles.Items.Count - 1;
                    if (currentIndex < lastIndex)
                    {
                        lvFiles.Items.Remove(selectedItem);
                        lvFiles.Items.Insert(lastIndex, selectedItem);
                    }
                }
            }
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            if (lvFiles.SelectedItems.Count > 0)
            {
                foreach (ListViewItem selectedItem in lvFiles.SelectedItems)
                {
                    lvFiles.Items.Remove(selectedItem);
                }
            }
        }

        [STAThread]
        private void BtnMerge_Click(object sender, EventArgs e)
        {
            try
            {
                // Get selected break option
                WdBreakType breakType = GetBreakTypeFromSelectedOption(cbBreakOptions.Text);

                // Set up progress bar
                int totalDocuments = lvFiles.Items.Count;
                int currentDocument = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = totalDocuments;
                progressBar1.Value = 0;
                progressBar1.Visible = true;


                // Merge documents
                foreach (ListViewItem item in lvFiles.Items)
                {
                    currentDocument++;

                    // Update progress bar
                    progressBar1.Value = currentDocument;

                    // Get file path
                    string filePath = item.SubItems[2].Text;


                    // Check if document is protected
                    if (IsDocumentProtected(filePath))
                    {
                        // Prompt for password
                        password = GetPasswordForProtectedDocument(filePath);
                        if (string.IsNullOrEmpty(password))
                        {
                            // User cancelled entering password, skip this document
                            continue;
                        }

                        // Open protected document with password
                        Document doc = OpenDocument(filePath, password);

                        if (doc == null)
                        {
                            Misc.ShowError($"Failed to open protected document: {filePath}");
                            continue;
                        }

                        // Add document to merged documents list
                        mergedDocuments.Add(doc);
                    }
                    else
                    {
                        // Open document
                        Document doc = OpenDocument(filePath);
                        if (doc == null)
                        {
                            Misc.ShowError($"Failed to open document: {filePath}");

                            continue;
                        }

                        // Add document to merged documents list
                        mergedDocuments.Add(doc);
                    }

                    // Update progress bar
                    progressBar1.Value = currentDocument;
                }

                // Merge the documents
                Document mergedDoc = MergeDocuments(mergedDocuments, breakType);
                if (mergedDoc != null)
                {
                    // Save the merged document
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Word Document (*.docx)|*.docx";
                    saveFileDialog.Title = "Save Merged File to:";
                    saveFileDialog.FileName = "_Merged";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string outputFile = saveFileDialog.FileName;
                        mergedDoc.SaveAs2(outputFile);
                        mergedDoc.Close();
                        Misc.ShowInformation("Merge completed successfully!");
                    }
                }

                // Hide progress bar
                progressBar1.Visible = false;
            }
            catch (Exception ex)
            {
                Misc.ShowError($"An error occurred during merging: {ex.Message}");
            }
            finally
            {
                // Release the Word application
                if (Globals.ThisAddIn.Application != null)
                {
                    Globals.ThisAddIn.Application.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Globals.ThisAddIn.Application);
                    Globals.ThisAddIn.Application = null;
                }
            }
        }


        private bool IsDocumentProtected(string filePath)
        {
            bool isProtected = false;
            Document doc = null;

            try
            {
                doc = Globals.ThisAddIn.Application.Documents.Open(filePath);
                isProtected = doc.ProtectionType != WdProtectionType.wdNoProtection;
            }
            catch (Exception ex)
            {
                Misc.ShowError($"Error checking document protection: {ex.Message}");
            }

            return isProtected;
        }

        private string GetPasswordForProtectedDocument(string filePath)
        {
            using (var passwordDialog = new PasswordDialog())
            {
                if (passwordDialog.ShowDialog() == DialogResult.OK)
                {
                    return passwordDialog.Password;
                }
                else
                {
                    return string.Empty;
                }
            }
        }

        private Document OpenDocument(string filePath, string password = null)
        {
            object missing = Type.Missing;
            object readOnly = false;
            object isVisible = true;
            object confirmConversions = false;
            object addToRecentFiles = false;
            object passwordObj = password;
            object noEncodingDialog = true; // Add this line to prevent the encoding dialog

            try
            {
                if (password == null)
                {
                    return Globals.ThisAddIn.Application.Documents.Open(filePath, ref confirmConversions, ref readOnly, ref addToRecentFiles,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref isVisible, ref missing, ref missing, ref noEncodingDialog, ref missing);
                }
                else
                {
                    return Globals.ThisAddIn.Application.Documents.Open(filePath, ref confirmConversions, ref readOnly, ref addToRecentFiles,
                        ref passwordObj, ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref isVisible, ref missing, ref missing, ref noEncodingDialog, ref missing);
                }
            }
            catch (Exception ex)
            {
                // Handle or log the exception
                Misc.ShowError($"Error occurred while opening the document: {ex.Message}");
                return null;
            }
        }

        private Document MergeDocuments(List<Document> documents, WdBreakType breakType)
        {
            // Check if there are any documents to merge
            if (documents.Count == 0)
            {
                Misc.ShowInformation("No documents to merge.");
                return null;
            }

            try
            {
                // Get the first document from the list
                Document firstDoc = documents[0];
                documents.RemoveAt(0);

                // Activate the first document
                firstDoc.Activate();

                if (IsDocumentProtected(firstDoc.FullName))
                {
                    firstDoc.Unprotect(password);
                }

                firstDoc.TrackRevisions = false;

                // Loop through each document in the list
                var i = 0;
                foreach (Document doc in documents)
                {
                    doc.TrackRevisions = false;
                    // Copy the content of the current document and paste it into the merged document
                    Range sourceRange = doc.Content;
                    Range targetRange = firstDoc.Content;

                    sourceRange.Copy();
                    targetRange.Characters.Last.FormattedText = sourceRange.FormattedText;

                    // Insert a section break (specified by breakType) between documents
                    if (i != documents.Count - 1) targetRange.Characters.Last.InsertBreak(breakType);
                    i++;
                    doc.Close(false);
                }

                return firstDoc;
            }
            catch (Exception ex)
            {
                Misc.ShowInformation($"Error merging documents: {ex.Message}");
                return null;
            }
        }

        // Sort the ListView by "File Name" using alphanumeric sorting
        private void SortListViewByFileName(int columnIndex)
        {
            ListViewItem[] items = new ListViewItem[lvFiles.Items.Count];
            lvFiles.Items.CopyTo(items, 0);

            Array.Sort(items, new ListViewItemComparer(columnIndex));

            lvFiles.Items.Clear();
            lvFiles.Items.AddRange(items);
        }

        // ListViewItemComparer class for alphanumeric sorting by "File Name"
        public class ListViewItemComparer : IComparer<ListViewItem>
        {
            private readonly int columnIndex;

            public ListViewItemComparer(int columnIndex)
            {
                this.columnIndex = columnIndex;
            }

            public int Compare(ListViewItem x, ListViewItem y)
            {
                string nameX = x.SubItems[columnIndex].Text;
                string nameY = y.SubItems[columnIndex].Text;

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
        private WdBreakType GetBreakTypeFromSelectedOption(string selectedOption)
        {
            switch (selectedOption)
            {
                case "Section Break":
                    return WdBreakType.wdSectionBreakNextPage;
                case "Page Break":
                    return WdBreakType.wdPageBreak;
                case "Line Break":
                    return WdBreakType.wdLineBreak;
                default:
                    return WdBreakType.wdSectionBreakNextPage;
            }
        }
    }
}
