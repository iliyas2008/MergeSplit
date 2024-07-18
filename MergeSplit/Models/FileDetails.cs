using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeSplit.Models
{
    public class FileDetails : INotifyPropertyChanged
    {
        private string fileName;
        public string FileName
        {
            get { return fileName; }
            set
            {
                fileName = value;
                OnPropertyChanged(nameof(FileName));
            }
        }
        private string fileFullName;
        public string FileFullName
        {
            get { return fileFullName; }
            set
            {
                fileFullName = value;
                OnPropertyChanged(nameof(FileFullName));
            }
        }
        private string lastModified;
        public string LastModified
        {
            get { return lastModified; }
            set
            {
                lastModified = value;
                OnPropertyChanged(nameof(LastModified));
            }
        }
        
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
