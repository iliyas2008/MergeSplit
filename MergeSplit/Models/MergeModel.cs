using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeSplit.Models
{
    internal class MergeModel
    {
        public ObservableCollection<FileDetails> MergeFiles { get; set; }
        public bool AcceptRevisions { get; set; }
        public int BreakOptionsIndex { get; set; }
    }
}
