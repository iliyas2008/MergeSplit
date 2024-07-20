using System.Windows;
using System.Windows.Controls;

namespace MergeSplit
{
    public partial class MainWindow : Window
    {
        private bool canNavigate = true;
        public MainWindow()
        {
            InitializeComponent();
            MainTabControl.SelectionChanged += MainTabControl_SelectionChanged;
            Tab1Frame.Navigating += Tab1Frame_Navigating;
        }
        private void MainTabControl_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (MainTabControl.SelectedIndex == 0) // Tab 1 selected
            {
                // Create an instance of MergeUserControl and set it as the content of Tab1ContentControl
                Tab1Frame.Navigate(new MergeWindow());
            }
            else
            {
                // Clear content if Tab 1 is not selected
                Tab1Frame.Content = null;
            }
        }
        private void Tab1Frame_Navigating(object sender, System.Windows.Navigation.NavigatingCancelEventArgs e)
        {
            if (!canNavigate)
            {
                e.Cancel = true;
            }
        }

        public void ListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Disable navigation to new pages
            canNavigate = false;
        }

        public void EnableNavigation()
        {
            // Enable navigation to new pages
            canNavigate = true;
        }
    }
}
