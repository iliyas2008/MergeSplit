﻿using MergeSplit.Models;
using MergeSplit.ViewModels;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace MergeSplit.Views
{
    /// <summary>
    /// Interaction logic for MergeUserControl.xaml
    /// </summary>
    public partial class MergeUserControl : UserControl
    {
        private MergeViewModel _viewModel;
        public MergeUserControl()
        {
            InitializeComponent();
            _viewModel = new MergeViewModel();
            DataContext = _viewModel;
        }
        private void GridViewColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            var headerClicked = e.OriginalSource as GridViewColumnHeader;
            if (headerClicked != null)
            {
                int columnIndex = Convert.ToInt32(headerClicked.Tag);
                Sort(columnIndex);
            }
        }
        private void Sort(int columnIndex)
        {
            var sortedFiles = _viewModel.Files.ToList();
            sortedFiles.Sort(new AlphanumericComparer(columnIndex));
            _viewModel.Files = new ObservableCollection<FileDetails>(sortedFiles);
        }
        public void ListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_viewModel != null)
            {
                _viewModel.SelectionChanged(lvFiles.SelectedItems);
            }

        }
    }
}
