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
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

namespace MergeSplit.ViewModels
{
    public class SplitViewModel : ViewModelBase
    {
        private bool _isDefaultPathChecked = true;
        private bool _isFilePathEnabled=false;
        private bool _isOpenDirEnabled=false;
        private ObservableCollection<string> _splitOptions;
        private string _selectedSplitOption;
        private string _prefix= "Chapter_";
        private string _filePath;
        private FileInfo fileInfo = null;
        private ObservableCollection<string> _previewItems;
        private ObservableCollection<FileDetails> _splitFiles;
        private WordprocessingDocument _doc = null;
        private bool _isNewgenChecked=true;
        private bool _isFMChecked = true;
        private bool _isIntroChecked = true;
        private bool _isBMChecked = true;
        private bool _isButtonEnabled = false;
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
        public ObservableCollection<string> SplitOptions
        {
            get => _splitOptions;
            set
            {
                _splitOptions = value;
                OnPropertyChanged(nameof(SplitOptions));
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

        public bool IsNewgenChecked
        {
            get => _isNewgenChecked;
            set
            {
                _isNewgenChecked = value;
                OnPropertyChanged(nameof(IsNewgenChecked));
            }
        }
        public bool IsFMChecked
        {
            get => _isFMChecked;
            set
            {
                _isFMChecked = value;
                OnPropertyChanged(nameof(IsFMChecked));
            }
        }

        public bool IsIntroChecked
        {
            get => _isIntroChecked;
            set
            {
                _isIntroChecked = value;
                OnPropertyChanged(nameof(IsIntroChecked));
            }
        }

        public bool IsBMChecked
        {
            get => _isBMChecked;
            set
            {
                _isBMChecked = value;
                OnPropertyChanged(nameof(IsBMChecked));
            }
        }
        public bool IsButtonEnabled
        {
            get => _isButtonEnabled;
            set
            {
                _isButtonEnabled = value;
                OnPropertyChanged(nameof(IsButtonEnabled));
            }
        }
        public ObservableCollection<string> PreviewItems
        {
            get => _previewItems;
            set
            {
                _previewItems = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsButtonEnabled));
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
            PreviewItems = new ObservableCollection<string>();
            SplitFiles = new ObservableCollection<FileDetails>();

            SplitOptions = new ObservableCollection<string> { "Section Break", "Page Break", "Heading 1 Style" };
            SelectedSplitOption = SplitOptions[0]; // Default selected option

        }
        public ClientModel ToClientModel()
        {
            return new ClientModel
            {
                IsNewgen = IsNewgenChecked,
                HasFM = IsFMChecked,
                HasIntro = IsIntroChecked,
                HasBM = IsBMChecked
            };
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
            switch (SelectedSplitOption.ToLower())
            {
                case "section break":
                    if (GetDocumentBreakOptions("section").SectionCount > 0)
                    {
                        Console.WriteLine($"There is {GetDocumentBreakOptions("section").SectionCount} Section Breaks");

                        if (IsNewgenChecked && !string.IsNullOrEmpty(FilePath))
                        {
                            GenerateSectionBreakPreview();
                        }
                        else
                        {
                            System.Windows.MessageBox.Show($"It is not Newgen Client.");
                        }

                    }
                    else
                    {
                        System.Windows.MessageBox.Show($"There is no Section Break.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    break;
                case "page break":
                    if (GetDocumentBreakOptions("page").PageBreakCount > 0)
                    {
                        Console.WriteLine($"There is {GetDocumentBreakOptions("page").PageBreakCount} Page Break");
                        if (IsNewgenChecked && !string.IsNullOrEmpty(FilePath))
                        {
                            GeneratePageBreakPreview();
                        }
                        else
                        {
                            System.Windows.MessageBox.Show($"It is not Newgen Client.");
                        }
                        
                    }
                    else
                    {
                        System.Windows.MessageBox.Show($"There is no Page Break.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    break;
                case "heading 1 style":
                    if (GetDocumentBreakOptions("heading").HeadingCount > 0)
                    {
                        Console.WriteLine($"There is {GetDocumentBreakOptions("heading").HeadingCount} Heading 1 Styles");
                        //GenerateHeading1StylePreview();
                    }
                    else
                    {
                        System.Windows.MessageBox.Show($"There is no Heading 1 Style.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    break;
                default:
                    if (GetDocumentBreakOptions("section").SectionCount > 0)
                    {
                        Console.WriteLine($"There is {GetDocumentBreakOptions("section").SectionCount} Section Breaks");

                        if (IsNewgenChecked && !string.IsNullOrEmpty(FilePath))
                        {
                            GenerateSectionBreakPreview();
                        }
                        else
                        {
                            System.Windows.MessageBox.Show($"It is not Newgen Client.");
                        }

                    }
                    else
                    {
                        System.Windows.MessageBox.Show($"There is no Section Break.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    break;
            }
            if (PreviewItems.Count > 0)
            {
                IsButtonEnabled = true;
            }
        }

        public class DocumentBreakOptions
        {
            public int SectionCount { get; set; }
            public int PageBreakCount { get; set; }
            public int HeadingCount { get; set; }
        }
        private DocumentBreakOptions GetDocumentBreakOptions(string breakString)
        {
            var result = new DocumentBreakOptions();

            if (string.IsNullOrWhiteSpace(FilePath) || !File.Exists(FilePath))
            {
                return result;
            }

            try
            {
                using (_doc = WordprocessingDocument.Open(FilePath, false))
                {
                    MainDocumentPart mainPart = _doc.MainDocumentPart;
                    var documentBody = mainPart.Document.Body;

                    if (breakString.Contains("section"))
                    {
                        // Count SectionProperties as section breaks
                        result.SectionCount = documentBody
                            .SelectMany(ox => ox.Descendants<SectionProperties>())
                            .Count();
                    }

                    if (breakString.Contains("page"))
                    {
                        // Count Page Breaks
                        result.PageBreakCount = documentBody.Elements<Paragraph>()
                            .SelectMany(p => p.Descendants<Run>())
                            .SelectMany(run => run.Descendants<Break>())
                            .Count(brk => brk.Type == BreakValues.Page);
                    }

                    if (breakString.Contains("heading"))
                    {
                        // Count Heading1 Paragraphs
                        result.HeadingCount = documentBody.Elements<Paragraph>()
                            .Count(p => p.ParagraphProperties?.ParagraphStyleId?.Val == "Heading1");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error processing the Word document: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return result;
        }

        private void GenerateSectionBreakPreview()
        {
            if (string.IsNullOrEmpty(FilePath)) return;

            int fmIndex = 1;
            int bmIndex = 1;
            int chapterIndex = 1;

            PreviewItems.Clear(); // Assuming PreviewItems is similar to lvPreview.Items

            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(FilePath, false))
                {
                    var documentBody = wordDoc.MainDocumentPart.Document.Body;
                    int sectionBreaks = documentBody.SelectMany(ox => ox.Descendants<SectionProperties>())
                            .Count();

                    for (int sectionNumber = 0; sectionNumber <= sectionBreaks; sectionNumber++)
                    {
                        string fileName = string.Empty;

                        if (IsFMChecked && sectionNumber == 0)
                        {
                            fileName = $"FM_{fmIndex}.docx";
                            fmIndex++;
                        }
                        else if (IsIntroChecked && sectionNumber == 1)
                        {
                            fileName = "Introduction.docx";
                        }
                        else if (IsBMChecked && sectionNumber == sectionBreaks)
                        {
                            fileName = $"BM_{bmIndex}.docx";
                            bmIndex++;
                        }
                        else
                        {
                            fileName = $"Chapter_{chapterIndex}.docx";
                            chapterIndex++;
                        }

                        PreviewItems.Add(fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Generate SectionBreak Preview: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void GeneratePageBreakPreview()
        {
            if (string.IsNullOrEmpty(FilePath)) return;

            int fmIndex = 1;
            int bmIndex = 1;
            int chapterIndex = 1; 
            
            PreviewItems.Clear();
            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(FilePath, false))
                {
                    var pageBreaks = wordDoc.MainDocumentPart.Document.Body.Elements<Paragraph>()
                                .SelectMany(p => p.Descendants<Run>())
                                .SelectMany(run => run.Descendants<Break>())
                                .Count(brk => brk.Type == BreakValues.Page);

                    for (int sectionNumber = 0; sectionNumber <= pageBreaks; sectionNumber++)
                    {
                        string fileName = string.Empty;

                        if (IsFMChecked && sectionNumber == 0)
                        {
                            fileName = $"FM_{fmIndex}.docx";
                            fmIndex++;
                        }
                        else if (IsIntroChecked && sectionNumber == 1)
                        {
                            fileName = "Introduction.docx";
                        }
                        else if (IsBMChecked && sectionNumber == pageBreaks)
                        {
                            fileName = $"BM_{bmIndex}.docx";
                            bmIndex++;
                        }
                        else
                        {
                            fileName = $"Chapter_{chapterIndex}.docx";
                            chapterIndex++;
                        }

                        PreviewItems.Add(fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Generate PageBreak Preview: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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
