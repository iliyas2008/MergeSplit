using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeSplit.ViewModels
{
    public class MainViewModel_Tabbed : ViewModelBase
    {
        public MainViewModel_Tabbed()
        {
            MergeViewModel = new MergeViewModel();
        }

        public MergeViewModel MergeViewModel { get; private set; }
    }
}
