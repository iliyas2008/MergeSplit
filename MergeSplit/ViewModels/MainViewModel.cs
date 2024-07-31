using MergeSplit.Commands;
using MergeSplit.Models;
using MergeSplit.Views;
using System;
using System.Collections.ObjectModel;
using System.Windows.Controls;
using System.Windows.Input;

namespace MergeSplit.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        
        private ObservableCollection<SidebarItem> _sidebarItems;
        private UserControl _currentView;
            public ICommand ChangeViewCommand { get; }

        public ObservableCollection<SidebarItem> SidebarItems
        {
            get => _sidebarItems;
            set
            {
                _sidebarItems = value;
                OnPropertyChanged();
            }
        }
        public UserControl CurrentView
            {
                get => _currentView;
                set => SetProperty(ref _currentView, value);
            }

            public MainViewModel()
            {
                SidebarItems = new ObservableCollection<SidebarItem>
            {
                new SidebarItem { Title = "Merge", ViewType = typeof(MergeUserControl) },
                new SidebarItem { Title = "Split", ViewType = typeof(SplitUserControl) },
                // Add more views as needed
            };
            // Set default view to MergeUserControl
            CurrentView = new MergeUserControl();
            SidebarItems[0].IsSelected = true;
            ChangeViewCommand = new DelegateCommand(ChangeView);
            }

            private void ChangeView(object parameter)
            {
                if (parameter is SidebarItem sidebarItem)
                {

                    foreach (var item in SidebarItems)
                    {
                        item.IsSelected = false;
                    }

                    sidebarItem.IsSelected = true;

                    CurrentView = (UserControl)Activator.CreateInstance(sidebarItem.ViewType);
                }
            }
        }
}
