using System.Collections.ObjectModel;
using System.ComponentModel;

namespace MergeSplit.Models
{
    public class MergeModel : INotifyPropertyChanged
    {
        private ObservableCollection<FileDetails> fileInfos;
        private bool acceptRevisions;
        private int breakOptionsIndex;

        public ObservableCollection<FileDetails> FileInfos
        {
            get => fileInfos;
            set
            {
                fileInfos = value;
                OnPropertyChanged(nameof(FileInfos));
            }
        }

        public bool AcceptRevisions
        {
            get => acceptRevisions;
            set
            {
                acceptRevisions = value;
                OnPropertyChanged(nameof(AcceptRevisions));
            }
        }

        public int BreakOptionsIndex
        {
            get => breakOptionsIndex;
            set
            {
                breakOptionsIndex = value;
                OnPropertyChanged(nameof(BreakOptionsIndex));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
