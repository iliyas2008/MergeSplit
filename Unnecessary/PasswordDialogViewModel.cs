using System;
using System.ComponentModel;
using System.Windows.Input;

namespace MergeSplit.ViewModels
{
    public class PasswordDialogViewModel : INotifyPropertyChanged
    {
        private string _password;
        public string Password
        {
            get { return _password; }
            set
            {
                if (_password != value)
                {
                    _password = value;
                    OnPropertyChanged(nameof(Password));
                }
            }
        }
        
        public ICommand OKCommand { get; private set; }
        public ICommand CancelCommand { get; private set; }
        public Action<bool> CloseAction { get; set; }
        public PasswordDialogViewModel()
        {
            OKCommand = new RelayCommand(OK);
            CancelCommand = new RelayCommand(Cancel);
        }

        private void OK(object obj)
        {
            CloseAction?.Invoke(true);
        }

        private void Cancel(object obj)
        {
            Password = string.Empty;
            CloseAction?.Invoke(false);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }

}