using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MergeSplit.Models;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;

namespace MergeSplit.ViewModels
{
    public class SplitViewModel : ViewModelBase
    {
        private bool _isDefaultPathChecked = true;
        private bool _isFilePathEnabled=false;
        private bool _isOpenDirEnabled=false;
        private string _selectedSplitOption;
        private string _prefix= "Chapter_";
        private string _filePath;
        private FileInfo fileInfo = null;
        private ObservableCollection<FileDetails> _previewItems;
        private ObservableCollection<FileDetails> _splitFiles;
        private WordprocessingDocument _doc = null;
        public bool IsDefaultPathChecked
        {
            get => _isDefaultPathChecked;
            set
            {
                if (_isDefaultPathChecked != value)
                {
                    _isDefaultPathChecked = value;
                    OnPropertyChanged();
                    UpdateControls();
                }
            }
        }

        public bool IsFilePathEnabled
        {
            get => _isFilePathEnabled;
            set
            {
                _isFilePathEnabled = value;
                OnPropertyChanged();
            }
        }
        public string FilePath
        {
            get => _filePath;
            set
            {
                _filePath = value;
                OnPropertyChanged();
            }
        }
        public bool IsOpenDirEnabled
        {
            get => _isOpenDirEnabled;
            set
            {
                _isOpenDirEnabled = value;
                OnPropertyChanged();
            }
        }
        public string Prefix
        {
            get => _prefix;
            set
            {
                _prefix = value;
                OnPropertyChanged();
            }
        }
        public string SelectedSplitOption
        {
            get => _selectedSplitOption;
            set
            {
                if (_selectedSplitOption != value)
                {
                    _selectedSplitOption = value;
                    OnPropertyChanged();
                }
            }
        }
        public ObservableCollection<FileDetails> PreviewItems
        {
            get => _previewItems;
            set
            {
                _previewItems = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<FileDetails> SplitFiles
        {
            get => _splitFiles;
            set
            {
                _splitFiles = value;
                OnPropertyChanged();
            }
        }
        // Event to request window close
        public event Action RequestClose;

        public ICommand SetDefaultPathCommand { get; }
        public ICommand CloseCommand { get; }

        public ICommand ResetCommand { get; }
        public ICommand OpenDirCommand { get; }
        public ICommand ChangePrefixCommand { get; }

        public SplitViewModel()
        {
            CloseCommand = new RelayCommand(param => OnClose());
            ResetCommand = new RelayCommand(param => Reset());
            OpenDirCommand = new RelayCommand(param => OpenDirectory());
            ChangePrefixCommand = new RelayCommand(param => ChangePrefix());
            PreviewItems = new ObservableCollection<FileDetails>();
            SplitFiles = new ObservableCollection<FileDetails>();

            SelectedSplitOption = "Section Break";
        }

        public void HandleFileDrop(string filePath)
        {
            string fileExtension = Path.GetExtension(filePath).ToLower();
            if (fileExtension.Equals(".docx") || fileExtension.Equals(".doc"))
            {
                FilePath = filePath;
                Console.WriteLine(FilePath);
                ConvertToOpenXmlDocument(filePath);
            }
            else
            {
                System.Windows.MessageBox.Show("Please drop a valid Word file (.docx or .doc).", "Invalid File", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ConvertToOpenXmlDocument(string filePath)
        {
            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
                {
                    fileInfo = new FileInfo(filePath);
                    Console.WriteLine($"Document Title: {fileInfo.Name}");
                    /*var docProps = wordDoc.PackageProperties;
                    if (docProps != null)
                    {
                        var title = docProps.Title;
                        Console.WriteLine($"Document Title: {title}");
                    }*/
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error processing the Word document: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void ChangePrefix()
        {
            if (string.IsNullOrEmpty(SelectedSplitOption))
            {
                System.Windows.MessageBox.Show($"Please select a split option.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            switch (SelectedSplitOption)
            {
                case "Section Break":
                    if (DocumentHasBreakOptions("section"))
                    {
                        System.Windows.Forms.MessageBox.Show("There is a Section Break");
                        //GenerateSectionBreakPreview();
                    }
                    else
                    {
                        System.Windows.MessageBox.Show($"There is no Section Break.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    break;
                case "Page Break":
                    if (DocumentHasBreakOptions("page"))
                    {
                        System.Windows.Forms.MessageBox.Show("There is a Page Break");//GeneratePageBreakPreview();
                    }
                    else
                    {
                        System.Windows.MessageBox.Show($"There is no Page Break.", "Error", MessageBoxButton.OK, MessageBoxImage.Error); 
                    }
                    break;
                case "Heading 1 Style":
                    if (DocumentHasBreakOptions("heading"))
                    {
                        System.Windows.Forms.MessageBox.Show("There is a Heading 1 Style");//GenerateHeading1StylePreview();
                    }
                    else
                    {
                        System.Windows.MessageBox.Show($"There is no Heading 1 Style.", "Error", MessageBoxButton.OK, MessageBoxImage.Error); 
                    }
                    break;
                default:
                    System.Windows.MessageBox.Show($"Unknown split option selected.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    break;
            }

            // Enable the split button if there are items to split
            if (SplitFiles.Count > 0)
            {
                // Notify the view to enable the split button
                OnPropertyChanged(nameof(SplitFiles));
            }
        }

        private bool DocumentHasBreakOptions(string breakString)
        {
            if (string.IsNullOrWhiteSpace(FilePath) || !File.Exists(FilePath))
            {
                return false;
            }

            try
            {
                using (_doc = WordprocessingDocument.Open(FilePath, false))
                {
                    MainDocumentPart mainPart = _doc.MainDocumentPart;

                    if (breakString.Contains("section"))
                    {
                        var sections = mainPart.Document.Body.Elements<SectionProperties>();
                        return sections.Any();
                    }
                    else if (breakString.Contains("page"))
                    {
                        var paragraphs = mainPart.Document.Body.Elements<Paragraph>();
                        return paragraphs.Any(paragraph => paragraph.Descendants<Run>()
                            .Any(run => run.Descendants<Break>().Any(brk => brk.Type == BreakValues.Page)));
                    }
                    else if (breakString.Contains("heading"))
                    {
                        var headingParagraphs = mainPart.Document.Body.Elements<Paragraph>()
                            .Where(p => p.ParagraphProperties?.ParagraphStyleId?.Val == "Heading1");
                        return headingParagraphs.Any();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error processing the Word document: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return false;
        }
       
        private void UpdateControls()
        {
            IsFilePathEnabled = !IsDefaultPathChecked;
            IsOpenDirEnabled = !IsDefaultPathChecked;
        }
        
        private void Reset()
        {
            Prefix = "Chapter_";
            PreviewItems.Clear();
            SplitFiles.Clear();
        }
        private void OpenDirectory()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                DialogResult result = dialog.ShowDialog();
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    FilePath = dialog.SelectedPath;
                }
            }
        }
        private void OnClose()
        {
            RequestClose?.Invoke();
        }
    }
}
