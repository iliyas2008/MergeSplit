using System.ComponentModel;

namespace MergeSplit.Models
{
    public class ClientModel : INotifyPropertyChanged
    {
        private bool isNewgen;
        private bool hasFM;
        private bool hasIntro;
        private bool hasBM;
        public bool IsNewgen
        {
            get { return isNewgen; }
            set
            {
                isNewgen = value;
                OnPropertyChanged(nameof(isNewgen));
            }
        }
        public bool HasFM
        {
            get { return hasFM; }
            set
            {
                hasFM = value;
                OnPropertyChanged(nameof(hasFM));
            }
        }
        public bool HasIntro
        {
            get { return hasIntro; }
            set
            {
                isNewgen = value;
                OnPropertyChanged(nameof(hasIntro));
            }
        }
        public bool HasBM
        {
            get { return hasBM; }
            set
            {
                isNewgen = value;
                OnPropertyChanged(nameof(hasBM));
            }
        }
       
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
