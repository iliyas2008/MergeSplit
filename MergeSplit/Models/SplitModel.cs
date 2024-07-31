using System;
using System.Collections.ObjectModel;

namespace MergeSplit.Models
{
    public class SplitModel
    {
        private string _outputPath;
        private string _prefix;
        private string _breakOption;
        private bool _isNewgen;
        private bool _hasFM;
        private bool _hasBM;
        private bool _hasIntro;
        private ObservableCollection<FileDetails> _splitFiles;

        public SplitModel(string outputPath, string prefix, string breakOption, bool isNewgen, bool hasFM, bool hasBM, bool hasIntro, ObservableCollection<FileDetails> splitFiles)
        {
            _outputPath = outputPath ?? throw new ArgumentNullException(nameof(outputPath));
            _prefix = prefix ?? throw new ArgumentNullException(nameof(prefix));
            _breakOption = breakOption ?? throw new ArgumentNullException(nameof(breakOption));
            _isNewgen = isNewgen;
            _hasFM = hasFM;
            _hasBM = hasBM;
            _hasIntro = hasIntro;
            _splitFiles = splitFiles ?? throw new ArgumentNullException(nameof(splitFiles));
        }
    }
}
