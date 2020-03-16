using MVVMCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlockAttributePrj
{
    class EditAllBlockViewModel: ViewModelBase
    {
        public RelayCommand OKEditALlCmd { get; set; }
        public RelayCommand CancelEditAllCmd { get; set; }
        double textHeight;
        public double TextHeight
        {
            get
            {
                return textHeight;
            }
            set
            {
                if (textHeight!= value)
                {
                    textHeight = value;
                    RaisePropertyChanged("TextHeight");
                }
            }
        }
        bool invisible;
        public bool Invisible
        {
            get
            {
                return invisible;
            }
            set
            {
                if (invisible != value)
                {
                    invisible = value;
                    RaisePropertyChanged("Invisible");
                }
            }
        }
    }
}
