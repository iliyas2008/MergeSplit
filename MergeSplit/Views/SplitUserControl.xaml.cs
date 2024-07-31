using MergeSplit.ViewModels;
using System.Windows;
using System.Windows.Controls;

namespace MergeSplit.Views
{
    /// <summary>
    /// Interaction logic for SplitUserControl.xaml
    /// </summary>
    public partial class SplitUserControl : UserControl
    {
        private SplitViewModel _viewModel;
        public SplitUserControl()
        {
            InitializeComponent();
            _viewModel = new SplitViewModel();
            _viewModel.RequestClose += OnRequestClose;
            DataContext = _viewModel;
        }
        private void Border_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, true))
            {
                string[] fileNames = e.Data.GetData(DataFormats.FileDrop, true) as string[];
                if (fileNames != null && fileNames.Length == 1) // Only allow single file
                {
                    string fileExtension = System.IO.Path.GetExtension(fileNames[0]).ToLower();
                    if (fileExtension.Equals(".docx") || fileExtension.Equals(".doc"))
                    {
                        e.Effects = DragDropEffects.Copy;
                    }
                    else
                    {
                        e.Effects = DragDropEffects.None;
                    }
                }
                else
                {
                    e.Effects = DragDropEffects.None;
                }
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void Border_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] fileNames = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (fileNames != null && fileNames.Length == 1) // Only handle single file
                {
                    var viewModel = DataContext as SplitViewModel;
                    viewModel?.HandleFileDrop(fileNames[0]);
                }
                else
                {
                    MessageBox.Show("Please drop only one Word file at a time.", "Invalid Operation", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }
        private void OnRequestClose()
        {
            Window.GetWindow(this)?.Close();
        }
        
    }
}
