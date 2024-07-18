using MergeSplit.ViewModels;
using System;
using System.Windows;
using System.Windows.Controls;

namespace MergeSplit.Views
{
    /// <summary>
    /// Interaction logic for PasswordDialog.xaml
    /// </summary>
    public partial class PasswordDialog : Window
    {
        public event EventHandler<PasswordDialogResultEventArgs> PasswordDialogClosed;

        public PasswordDialog()
        {
            InitializeComponent();
        }

        private void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (DataContext is PasswordDialogViewModel viewModel)
            {
                viewModel.Password = passwordBox.Password;
            }
        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            PasswordDialogClosed?.Invoke(this, new PasswordDialogResultEventArgs(true, passwordBox.Password));
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            PasswordDialogClosed?.Invoke(this, new PasswordDialogResultEventArgs(false, string.Empty));
            Close();
        }
    }

    public class PasswordDialogResultEventArgs : EventArgs
    {
        public bool DialogResult { get; }
        public string Password { get; }

        public PasswordDialogResultEventArgs(bool dialogResult, string password)
        {
            DialogResult = dialogResult;
            Password = password;
        }
    }
}
