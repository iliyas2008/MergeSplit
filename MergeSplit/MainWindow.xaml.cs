using MergeSplit.ViewModels;
using System.Windows;

namespace MergeSplit
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel();
        }
    }
}
