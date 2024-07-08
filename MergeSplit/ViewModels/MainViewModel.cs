using MergeSplit.Models;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Data;
using System.Windows.Forms;
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
        private bool _isProgressBarVisible;
        public bool IsProgressBarVisible
        {
            get { return _isProgressBarVisible; }
            set
            {
                if (_isProgressBarVisible != value)
                {
                    _isProgressBarVisible = value;
                    OnPropertyChanged();
                }
            }
        }

        public int ProgressBarValue { get; private set; }

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
        public ICommand AddFolderCommand { get; private set; }
        public ICommand SelectionChangedCommand { get; }
        public ICommand MoveFirstCommand { get; }
        public ICommand MoveLastCommand { get; }
        public ICommand MoveUpCommand { get; }
        public ICommand MoveDownCommand { get; }
        public ICommand ClearListCommand { get; }
        public MainViewModel()
        {
            
            MergeModel = new MergeModel();

            Files = new ObservableCollection<FileDetails>();
            
            SelectedFiles = new ObservableCollection<FileDetails>();
            AddFilesCommand = new RelayCommand(param => AddFiles());
            AddFolderCommand = new RelayCommand(param => AddFolder());
            RemoveFilesCommand = new RelayCommand(param => RemoveFiles(), param => CanRemoveFiles());
            SelectionChangedCommand = new RelayCommand(param => SelectionChanged(param));
            MoveFirstCommand = new RelayCommand(param => MoveFirst(), param => CanMoveFirst());
            MoveLastCommand = new RelayCommand(param => MoveLast(), param => CanMoveLast());
            MoveUpCommand = new RelayCommand(param => MoveUp(), param => CanMoveUp());
            MoveDownCommand = new RelayCommand(param => MoveDown(), param => CanMoveDown());
            ClearListCommand = new RelayCommand(param => ClearList(), param => CanClearList());
        }
        private void AddFiles()
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog
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
                        LastModified = fileInfo.LastWriteTime.ToString()
                    });
                }
            }
            
        }
        private void AddFolder()
        {
            var folderBrowserDialog = new FolderBrowserDialog();

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                string folderPath = folderBrowserDialog.SelectedPath;
                AddFilesFromFolder(folderPath);
                SortListViewByFileName();
            }
        }
        private void AddFilesFromFolder(string folderPath)
        {
            string[] filesInFolder = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                .Where(file => !Path.GetFileName(file).StartsWith("~"))
                .ToArray();

            foreach (string filePath in filesInFolder)
            {
                FileInfo fileInfo = new FileInfo(filePath);

                if (!IsFileExtensionValid(fileInfo.Extension))
                    continue;

                FileDetails file = new FileDetails
                {
                    FileName = fileInfo.Name,
                    FileFullName = fileInfo.FullName,
                    LastModified = fileInfo.LastWriteTime.ToString()
                };

                Files.Add(file);
            }
        }
        private bool IsFileExtensionValid(string fileExtension)
        {
            string[] validExtensions = { ".doc", ".docx" };
            return validExtensions.Contains(fileExtension.ToLower());
        }
        private void SortListViewByFileName()
        {
            Files = new ObservableCollection<FileDetails>(Files.OrderBy(f => f.FileName));
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
