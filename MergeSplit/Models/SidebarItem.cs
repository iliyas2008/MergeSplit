using MergeSplit.ViewModels;
using System;

namespace MergeSplit.Models
{
    public class SidebarItem: ViewModelBase
    {
        public string Title { get; set; }
        public Type ViewType { get; set; }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }
    }
}
