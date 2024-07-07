using MergeSplit.ViewModels;
using System.Windows;
using System.Windows.Controls;

namespace MergeSplit
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        public void ListView_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var viewModel = DataContext as MainViewModel;
            if (viewModel != null)
            {
                viewModel.SelectionChanged(lvFiles.SelectedItems);
            }
        }

        private void chkAcceptRevisions_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void cbBreakOptions_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
    
}
