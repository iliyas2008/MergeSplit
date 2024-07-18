using MergeSplit.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Input;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Body = DocumentFormat.OpenXml.Wordprocessing.Body;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Style = DocumentFormat.OpenXml.Wordprocessing.Style;
using Endnote = DocumentFormat.OpenXml.Wordprocessing.Endnote;
using Footnote = DocumentFormat.OpenXml.Wordprocessing.Footnote;
using Comment = DocumentFormat.OpenXml.Wordprocessing.Comment;
using MessageBox = System.Windows.MessageBox;

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
        public ICommand AddFolderCommand { get; private set; }
        public ICommand SelectionChangedCommand { get; }
        public ICommand MoveFirstCommand { get; }
        public ICommand MoveLastCommand { get; }
        public ICommand MoveUpCommand { get; }
        public ICommand MoveDownCommand { get; }
        public ICommand ClearListCommand { get; }
        public ICommand MergeCommand { get; }
        public MainViewModel()
        {
            Files = new ObservableCollection<FileDetails>();
            MergeModel = new MergeModel
            {
                FileInfos = Files,
                AcceptRevisions = false,
                BreakOptionsIndex = 0
            };
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
            MergeCommand = new RelayCommand(param => MergeDocs(MergeModel), param => CanMergeDocs());
        }
        public void MergeDocs(MergeModel mergeModel)
        {
            string mergedFileName = Path.GetTempFileName();
            if (mergeModel == null)
                return;

            try
            {
                using (WordprocessingDocument mergedDoc = WordprocessingDocument.Create(mergedFileName, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = mergedDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());

                    for (int i = 0; i < mergeModel.FileInfos.Count; i++)
                    {
                        var fileDetails = mergeModel.FileInfos[i];
                        string filePath = fileDetails.FileFullName;

                        using (WordprocessingDocument srcDoc = WordprocessingDocument.Open(filePath, false))
                        {
                            // Copy main document part content
                            foreach (var element in srcDoc.MainDocumentPart.Document.Body.Elements())
                            {
                                mainPart.Document.Body.AppendChild(element.CloneNode(true));
                            }

                            // Copy styles part
                            if (srcDoc.MainDocumentPart.StyleDefinitionsPart != null)
                            {
                                if (mainPart.StyleDefinitionsPart == null)
                                {
                                    mainPart.AddNewPart<StyleDefinitionsPart>();
                                    mainPart.StyleDefinitionsPart.FeedData(srcDoc.MainDocumentPart.StyleDefinitionsPart.GetStream());
                                }
                                else
                                {
                                    MergeStyles(srcDoc.MainDocumentPart.StyleDefinitionsPart, mainPart.StyleDefinitionsPart);
                                }
                            }

                            // Copy comments part
                            if (srcDoc.MainDocumentPart.WordprocessingCommentsPart != null)
                            {
                                if (mainPart.WordprocessingCommentsPart == null)
                                {
                                    mainPart.AddPart(srcDoc.MainDocumentPart.WordprocessingCommentsPart);
                                }
                                else
                                {
                                    MergeComments(srcDoc.MainDocumentPart.WordprocessingCommentsPart, mainPart.WordprocessingCommentsPart);
                                }
                            }

                            // Copy footnotes part
                            if (srcDoc.MainDocumentPart.FootnotesPart != null)
                            {
                                if (mainPart.FootnotesPart == null)
                                {
                                    mainPart.AddPart(srcDoc.MainDocumentPart.FootnotesPart);
                                }
                                else
                                {
                                    MergeFootnotes(srcDoc.MainDocumentPart.FootnotesPart, mainPart.FootnotesPart);
                                }
                            }

                            // Copy endnotes part
                            if (srcDoc.MainDocumentPart.EndnotesPart != null)
                            {
                                if (mainPart.EndnotesPart == null)
                                {
                                    mainPart.AddPart(srcDoc.MainDocumentPart.EndnotesPart);
                                }
                                else
                                {
                                    MergeEndnotes(srcDoc.MainDocumentPart.EndnotesPart, mainPart.EndnotesPart);
                                }
                            }

                            // Copy headers and footers
                            foreach (var headerPart in srcDoc.MainDocumentPart.HeaderParts)
                            {
                                mainPart.AddPart(headerPart);
                            }

                            foreach (var footerPart in srcDoc.MainDocumentPart.FooterParts)
                            {
                                mainPart.AddPart(footerPart);
                            }

                            if (mergeModel.AcceptRevisions)
                            {
                                AcceptAllRevisions(mergedDoc);
                                mergedDoc.MainDocumentPart.Document.Save();
                            }
                        }

                        // Insert break after each document except the last one
                        if (i < mergeModel.FileInfos.Count - 1)
                        {
                            InsertBreak(mainPart.Document.Body, mergeModel.BreakOptionsIndex);
                        }
                    }
                    mainPart.Document.Save();
                }

                System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog
                {
                    Filter = "Word Document (*.docx)|*.docx",
                    Title = "Save Merged File to:",
                    FileName = "_Merged"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string outputFile = saveFileDialog.FileName;
                    File.Copy(mergedFileName, outputFile, true);
                    MessageBox.Show("Merge completed successfully!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred during merging: {ex.Message}");
            }
            finally
            {
                if (File.Exists(mergedFileName))
                {
                    File.Delete(mergedFileName);
                }
            }
        }

        private void MergeStyles(StyleDefinitionsPart srcStylesPart, StyleDefinitionsPart destStylesPart)
        {
            var srcStyles = srcStylesPart.Styles;
            var destStyles = destStylesPart.Styles;

            foreach (var srcStyle in srcStyles.Elements<Style>())
            {
                if (!destStyles.Elements<Style>().Any(s => s.StyleId == srcStyle.StyleId))
                {
                    destStyles.AppendChild(srcStyle.CloneNode(true));
                }
            }
            destStyles.Save();
        }

        private void MergeComments(WordprocessingCommentsPart srcCommentsPart, WordprocessingCommentsPart destCommentsPart)
        {
            var srcComments = srcCommentsPart.Comments;
            var destComments = destCommentsPart.Comments;

            foreach (var srcComment in srcComments.Elements<Comment>())
            {
                if (!destComments.Elements<Comment>().Any(c => c.Id == srcComment.Id))
                {
                    destComments.AppendChild(srcComment.CloneNode(true));
                }
            }
            destComments.Save();
        }

        private void MergeFootnotes(FootnotesPart srcFootnotesPart, FootnotesPart destFootnotesPart)
        {
            var srcFootnotes = srcFootnotesPart.Footnotes;
            var destFootnotes = destFootnotesPart.Footnotes;

            foreach (var srcFootnote in srcFootnotes.Elements<Footnote>())
            {
                if (!destFootnotes.Elements<Footnote>().Any(f => f.Id == srcFootnote.Id))
                {
                    destFootnotes.AppendChild(srcFootnote.CloneNode(true));
                }
            }
            destFootnotes.Save();
        }

        private void MergeEndnotes(EndnotesPart srcEndnotesPart, EndnotesPart destEndnotesPart)
        {
            var srcEndnotes = srcEndnotesPart.Endnotes;
            var destEndnotes = destEndnotesPart.Endnotes;

            foreach (var srcEndnote in srcEndnotes.Elements<Endnote>())
            {
                if (!destEndnotes.Elements<Endnote>().Any(e => e.Id == srcEndnote.Id))
                {
                    destEndnotes.AppendChild(srcEndnote.CloneNode(true));
                }
            }
            destEndnotes.Save();
        }

        private void InsertBreak(Body body, int breakType)
        {
            switch (breakType)
            {
                case 0: // Section Break (Next Page)
                    Paragraph breakPara = new Paragraph();
                    ParagraphProperties breakParaProps = new ParagraphProperties();
                    SectionProperties sectProps1 = new SectionProperties();

                    // Insert a section break
                    sectProps1.Append(new SectionType() { Val = SectionMarkValues.NextPage });

                    breakParaProps.Append(sectProps1);
                    breakPara.Append(breakParaProps);

                    body.AppendChild(breakPara);
                    break;
                case 1: // Column Break
                    body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                    break;
                case 2: // Paragraph Break
                        // Create a new paragraph to simulate a paragraph break
                    body.AppendChild(new Paragraph());
                    break;
                default: // Default to Section Break (Next Page)
                    Paragraph defaultPara = new Paragraph();
                    ParagraphProperties defaultParaProps = new ParagraphProperties();
                    SectionProperties sectProps2 = new SectionProperties();

                    // Insert a section break
                    sectProps2.Append(new SectionType() { Val = SectionMarkValues.NextPage });

                    defaultParaProps.Append(sectProps2);
                    defaultPara.Append(defaultParaProps);

                    body.AppendChild(defaultPara);
                    break;
            }
        }

        public static void AcceptAllRevisions(WordprocessingDocument document)
        {
            var mainPart = document.MainDocumentPart;
            var body = mainPart.Document.Body;

            // Accept revisions in the main document
            AcceptRevisionsInElement(body);

            // Accept revisions in headers and footers
            foreach (var headerPart in mainPart.HeaderParts)
            {
                AcceptRevisionsInElement(headerPart.Header);
            }

            foreach (var footerPart in mainPart.FooterParts)
            {
                AcceptRevisionsInElement(footerPart.Footer);
            }

            // Accept revisions in footnotes
            var footnotesPart = mainPart.FootnotesPart;
            if (footnotesPart != null)
            {
                foreach (var footnote in footnotesPart.Footnotes.Elements<Footnote>())
                {
                    AcceptRevisionsInElement(footnote);
                }
            }

            // Accept revisions in endnotes
            var endnotesPart = mainPart.EndnotesPart;
            if (endnotesPart != null)
            {
                foreach (var endnote in endnotesPart.Endnotes.Elements<Endnote>())
                {
                    AcceptRevisionsInElement(endnote);
                }
            }

            // Save changes
            mainPart.Document.Save();
        }

        private static void AcceptRevisionsInElement(OpenXmlElement element)
        {
            // Handle the formatting changes.
            List<OpenXmlElement> changes = element.Descendants<ParagraphPropertiesChange>().Cast<OpenXmlElement>().ToList();
            changes.AddRange(element.Descendants<RunPropertiesChange>().Cast<OpenXmlElement>().ToList());
            changes.AddRange(element.Descendants<SectionPropertiesChange>().Cast<OpenXmlElement>().ToList());
            changes.AddRange(element.Descendants<TablePropertiesChange>().Cast<OpenXmlElement>().ToList());
            changes.AddRange(element.Descendants<TableGridChange>().Cast<OpenXmlElement>().ToList());
            changes.AddRange(element.Descendants<TableRowPropertiesChange>().Cast<OpenXmlElement>().ToList());
            changes.AddRange(element.Descendants<TableCellPropertiesChange>().Cast<OpenXmlElement>().ToList());


            foreach (OpenXmlElement change in changes)
            {
                change.Remove();
            }

            // Handle the deletions.
            List<OpenXmlElement> deletions = element
                .Descendants<Deleted>().Cast<OpenXmlElement>().ToList();

            deletions.AddRange(element.Descendants<DeletedRun>().Cast<OpenXmlElement>().ToList());

            deletions.AddRange(element.Descendants<DeletedMathControl>().Cast<OpenXmlElement>().ToList());

            foreach (OpenXmlElement deletion in deletions)
            {
                deletion.Remove();
            }

            // Handle the insertions.
            List<OpenXmlElement> insertions =
                element.Descendants<Inserted>().Cast<OpenXmlElement>().ToList();

            insertions.AddRange(element.Descendants<InsertedRun>().Cast<OpenXmlElement>().ToList());

            insertions.AddRange(element.Descendants<InsertedMathControl>().Cast<OpenXmlElement>().ToList());

            foreach (OpenXmlElement insertion in insertions)
            {
                // Found new content.
                // Promote them to the same level as node, and then delete the node.
                foreach (var run in insertion.Elements<Run>())
                {
                    if (run == insertion.FirstChild)
                    {
                        insertion.InsertAfterSelf(new Run(run.OuterXml));
                    }
                    else
                    {
                        OpenXmlElement nextSibling = insertion.NextSibling()!;
                        nextSibling.InsertAfterSelf(new Run(run.OuterXml));
                    }
                }

                insertion.RemoveAttribute("rsidR", "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
                insertion.RemoveAttribute("rsidRPr", "https://schemas.openxmlformats.org/wordprocessingml/2006/main");
                insertion.Remove();
            }
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
                //Files.Clear();
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
            var comparer = new AlphanumericComparer(0); // sorting by FileName (adjust index as needed)
            var sortedFiles = Files.OrderBy(f => f, comparer).ToList();
            Files.Clear();
            foreach (var file in sortedFiles)
            {
                Files.Add(file);
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
        private bool CanRemoveFiles() { return SelectedFiles.Count > 0;}
        private bool CanClearList() { return Files.Count > 0;}
        private bool CanMoveFirst() { return SelectedFiles.Count > 0; }
        private bool CanMoveLast() { return SelectedFiles.Count > 0; }
        private bool CanMoveUp() { return SelectedFiles.Count > 0 && SelectedFiles.Any(file => Files.IndexOf(file) > 0); }
        private bool CanMoveDown() { return SelectedFiles.Count > 0 && SelectedFiles.Any(file => Files.IndexOf(file) < Files.Count - 1); }
        private bool CanMergeDocs() { return Files.Count > 1; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
