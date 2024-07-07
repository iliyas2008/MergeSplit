using MergeSplit.Models;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;

namespace MergeSplit.ViewModels
{
    internal class MainViewModel : INotifyPropertyChanged
    {
        private ObservableCollection<FileDetails> _files;
        public ObservableCollection<FileDetails> Files
        {
            get { return _files; }
            set
            {
                _files = value;
                OnPropertyChanged();
            }
        }
        private MergeModel _mergeModel;
        public MergeModel MergeModel
        {
            get { return _mergeModel; }
            set
            {
                _mergeModel = value;
                OnPropertyChanged();
            }
        }
        private ObservableCollection<FileDetails> _selectedFiles;
        public ObservableCollection<FileDetails> SelectedFiles
        {
            get { return _selectedFiles; }
            set
            {
                _selectedFiles = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanRemoveFiles));
                OnPropertyChanged(nameof(CanMoveFirst));
                OnPropertyChanged(nameof(CanMoveLast));
                OnPropertyChanged(nameof(CanMoveUp));
                OnPropertyChanged(nameof(CanMoveDown));
                OnPropertyChanged(nameof(CanClearList));
            }
        }
        public ICommand AddFilesCommand { get; }
        public ICommand RemoveFilesCommand { get; }
        public ICommand SelectionChangedCommand { get; }
        public ICommand MoveFirstCommand { get; }
        public ICommand MoveLastCommand { get; }
        public ICommand MoveUpCommand { get; }
        public ICommand MoveDownCommand { get; }
        public ICommand ClearListCommand { get; }
        public ICommand MergeCommand { get; }
        public MainViewModel()
        {
            
            MergeModel = new MergeModel();

            Files = new ObservableCollection<FileDetails>();
                
            SelectedFiles = new ObservableCollection<FileDetails>();
            AddFilesCommand = new RelayCommand(param => AddFiles());
            RemoveFilesCommand = new RelayCommand(param => RemoveFiles(), param => CanRemoveFiles());
            SelectionChangedCommand = new RelayCommand(param => SelectionChanged(param));
            MoveFirstCommand = new RelayCommand(param => MoveFirst(), param => CanMoveFirst());
            MoveLastCommand = new RelayCommand(param => MoveLast(), param => CanMoveLast());
            MoveUpCommand = new RelayCommand(param => MoveUp(), param => CanMoveUp());
            MoveDownCommand = new RelayCommand(param => MoveDown(), param => CanMoveDown());
            ClearListCommand = new RelayCommand(param => ClearList(), param => CanClearList());
            MergeCommand = new RelayCommand(param => MergeDocuments(), param => CanMergeDocuments());
        }
        private void AddFiles()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Word Documents|*.doc;*.docx"
            };

            bool? result = openFileDialog.ShowDialog();

            if (result == true)
            {
                Files.Clear();
                foreach (string filePath in openFileDialog.FileNames)
                {
                    var fileInfo = new FileInfo(filePath);
                    Files.Add(new FileDetails
                    {
                        FileName = fileInfo.Name,
                        FileFullName = fileInfo.FullName,
                        LastModified = fileInfo.LastWriteTime
                    });
                }
            }
        }
        private void MoveFirst()
        {
            var itemsToMove = SelectedFiles.ToList();
            foreach (var item in itemsToMove)
            {
                Files.Remove(item);
            }
            for (int i = itemsToMove.Count - 1; i >= 0; i--)
            {
                Files.Insert(0, itemsToMove[i]);
            }
            SelectedFiles.Clear();
        }
        private void MoveLast()
        {
            var itemsToMove = SelectedFiles.ToList();
            foreach (var item in itemsToMove)
            {
                Files.Remove(item);
            }
            foreach (var item in itemsToMove)
            {
                Files.Add(item);
            }
            SelectedFiles.Clear();
        }
        private void MoveUp()
        {
            var itemsToMove = SelectedFiles.ToList();
            foreach (var item in itemsToMove)
            {
                int index = Files.IndexOf(item);
                if (index > 0)
                {
                    Files.Move(index, index - 1);
                }
            }
        }

        private void MoveDown()
        {
            var itemsToMove = SelectedFiles.ToList();
            for (int i = itemsToMove.Count - 1; i >= 0; i--)
            {
                var item = itemsToMove[i];
                int index = Files.IndexOf(item);
                if (index < Files.Count - 1)
                {
                    Files.Move(index, index + 1);
                }
            }
        }
        private void RemoveFiles()
        {
            foreach (var file in SelectedFiles.ToList())
            {
                Files.Remove(file);
            }
            SelectedFiles.Clear();
        }
        public void SelectionChanged(object parameter)
        {
            SelectedFiles.Clear();

            if (parameter is IList selectedItems)
            {
                foreach (FileDetails file in selectedItems)
                {
                    SelectedFiles.Add(file);
                }
            }
        }
        private void ClearList()
        {
            Files.Clear();
            SelectedFiles.Clear();
        }
        private void MergeDocuments()
        {
            Microsoft.Office.Interop.Word.Application wordApp = null;
            Document mergedDoc = null;

            try
            {
                wordApp = new Microsoft.Office.Interop.Word.Application();
                // Create a new document to merge into
                mergedDoc = wordApp.Documents.Add();

                foreach (var file in MergeModel.MergeFiles)
                {
                    if (File.Exists(file.FileFullName))
                    {
                        // Open each document from the list
                        Document docToMerge = wordApp.Documents.Open(file.FileFullName);

                        // Merge the documents
                        if (mergedDoc != null)
                        {
                            if (MergeModel.AcceptRevisions)
                            {
                                // Accept all revisions in the document to be merged
                                docToMerge.Revisions.AcceptAll();
                            }

                            // Copy content from docToMerge to mergedDoc
                            Range rangeToCopy = docToMerge.Content;
                            rangeToCopy.Copy();
                            Range rangeToPaste = mergedDoc.Content;
                            rangeToPaste.Collapse(WdCollapseDirection.wdCollapseEnd);

                            // Apply break options if specified
                            switch (MergeModel.BreakOptionsIndex)
                            {
                                case 1: // Section Break
                                    rangeToPaste.InsertBreak(WdBreakType.wdSectionBreakNextPage);
                                    break;
                                case 2: // Page Break
                                    rangeToPaste.InsertBreak(WdBreakType.wdPageBreak); 
                                    break;
                                default: // None or default
                                    break;
                            }

                            // Paste the copied content
                            rangeToPaste.Paste();
                        }

                        // Close the document without saving changes
                        docToMerge.Close(WdSaveOptions.wdDoNotSaveChanges);
                    }
                }

                // Show the merged document
                wordApp.Visible = true;
            }
            catch (Exception ex)
            {
                // Handle exceptions
                MessageBox.Show($"Error merging documents: {ex.Message}");
            }
            finally
            {
                // Clean up resources
                if (mergedDoc != null) Marshal.ReleaseComObject(mergedDoc);
                if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            }
        }

        public bool CanMergeDocuments()
        {
            return MergeModel.MergeFiles.Count > 0;
        }
        private bool CanRemoveFiles()
        {
            return SelectedFiles.Count > 0;
        }
        public bool CanClearList()
        {
            return Files.Count > 0;
        }
        private bool CanMoveFirst()
        {
            return SelectedFiles.Count > 0;
        }
        private bool CanMoveLast()
        {
            return SelectedFiles.Count > 0;
        }
        public bool CanMoveUp()
        { 
            return SelectedFiles.Count > 0 && SelectedFiles.Any(file => Files.IndexOf(file) > 0);
        }
        public bool CanMoveDown() 
        { 
            return SelectedFiles.Count > 0 && SelectedFiles.Any(file => Files.IndexOf(file) < Files.Count - 1); 
        }
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
