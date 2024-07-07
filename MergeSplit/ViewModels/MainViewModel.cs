using MergeSplit.Models;
using Microsoft.Win32;
using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
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
        private ObservableCollection<FileDetails> _selectedFiles;
        public ObservableCollection<FileDetails> SelectedFiles
        {
            get { return _selectedFiles; }
            set
            {
                _selectedFiles = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CanRemoveFiles));
            }
        }
        public ICommand AddFilesCommand { get; }
        public ICommand RemoveFilesCommand { get; }
        public ICommand SelectionChangedCommand { get; }
        public MainViewModel()
        {
            Files = new ObservableCollection<FileDetails>();
            SelectedFiles = new ObservableCollection<FileDetails>();
            AddFilesCommand = new RelayCommand(param => AddFiles());
            RemoveFilesCommand = new RelayCommand(param => RemoveFiles(), param => CanRemoveFiles());
            SelectionChangedCommand = new RelayCommand(param => SelectionChanged(param));

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
        private bool CanRemoveFiles()
        {
            return SelectedFiles.Count > 0;
        }
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
