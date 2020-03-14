using Autodesk.AutoCAD.DatabaseServices;
using MVVMCore;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlockAttributePrj
{
    public class BlockAttribute
    {
        private bool bCheckInvisible = true;
        public bool CheckInvisible
        {
            get
            {
                return bCheckInvisible;
            }
            set
            {
                if (bCheckInvisible != value)
                {
                    bCheckInvisible = value;
                }
            }
        }
        private string _AtrributeTag="";
        public string AtrributeTag
        {
            get
            {
                return _AtrributeTag;
            }
            set
            {
                if (_AtrributeTag != value)
                {
                    _AtrributeTag = value;
                }
            }
        }
        private string _Promt="";
        public string Promt
        {
            get
            {
                return _Promt;
            }
            set
            {
                if (_Promt != value)
                {
                    _Promt = value;
                }
            }
        }
        private string _Value= "";
        public string Value
        {
            get
            {
                return _Value;
            }
            set
            {
                if (_Value != value)
                {
                    _Value = value;
                }
            }
        }
        private double _nTextHeight;
        public double TextHeight
        {
            get
            {
                return _nTextHeight;
            }
            set
            {
                if (_nTextHeight != value)
                {
                    _nTextHeight = value;
                }
            }
        }
        private string _TextStyle = "";
        public string TextStyle
        {
            get
            {
                return _TextStyle;
            }
            set
            {
                if (_TextStyle != value)
                {
                    _TextStyle = value;
                }
            }
        }
        ObservableCollection<string> _TextStyles;
        public ObservableCollection<string> TextStyles
        {
            get
            {
                return _TextStyles;
            }
            set
            {
                if (_TextStyles != value)
                {
                    _TextStyles = value;
                }
            }
        }
        public ObjectId AtrrId;
    }
    public class BlockReferenceProperties
    {
        public BlockReferenceProperties()
        {
            _blkAtrribute = new ObservableCollection<BlockAttribute>();
        }
        ObjectId _BlockId = ObjectId.Null;
        public ObjectId BlockId
        {
            get
            {
                return _BlockId;
            }
            set
            {
                if (_BlockId != value)
                {
                    _BlockId = value;
                }
            }
        }
        public string _BlockName = string.Empty;
        public string BlockName
        {
            get
            {
                return _BlockName;
            }
            set
            {
                if (_BlockName != value)
                {
                    _BlockName = value;
                }
            }
        }
        private BlockAttribute _Attributes ;
        public BlockAttribute Attributes
        {
            get
            {
                return _Attributes;
            }
            set
            {
                if (value!= _Attributes)
                {
                    _Attributes = value;
                }
            }
        }
        private ObservableCollection<BlockAttribute> _blkAtrribute;
        public ObservableCollection<BlockAttribute> blkAttribute
        {
            get
            {
                return _blkAtrribute;
            }
            set
            {
                if (_blkAtrribute!= value)
                {
                    _blkAtrribute = value;
                }
            }
        }
       
    }
    public class ModifyBlockViewModel:ViewModelBase
    {
        public ModifyBlockViewModel()
        {
            _blkProperties = new ObservableCollection<BlockReferenceProperties>();
            _blkAttributes = new ObservableCollection<BlockAttribute>();
        }
        public RelayCommand OKCmd { get; set; }
        public RelayCommand CancelCmd { get; set; }
        public RelayCommand UncheckAllCmd { get; set; }
        public RelayCommand CheckAllCmd { get; set; }
        private ObservableCollection <BlockReferenceProperties> _blkProperties;
        public ObservableCollection <BlockReferenceProperties> BlkProperties
        {
            get
            {
                return _blkProperties;
            }
            set
            {
                if (value != _blkProperties)
                {
                    _blkProperties = value;
                    RaisePropertyChanged("BlkProperties");
                }
            }
        }
        private BlockReferenceProperties _blkSelected;
        public BlockReferenceProperties blkSelected
        {
            get
            {
                return _blkSelected;
            }
            set
            {
                if (_blkSelected != value)
                {
                    _blkSelected = value;
                    RaisePropertyChanged("blkSelected");
                }
            }
        }
        private BlockAttribute _blkAttributeSelected;
        public BlockAttribute blkAttributeSelected
        {
            get
            {
                return _blkAttributeSelected;
            }
            set
            {
                if (_blkAttributeSelected != value)
                {
                    _blkAttributeSelected = value;
                    RaisePropertyChanged("blkAttributeSelected");
                }
            }
        }
        private ObservableCollection<BlockAttribute> _blkAttributes;
        public ObservableCollection<BlockAttribute> blkAttributes
        {
            get
            {
                return _blkAttributes;
            }
            set
            {
                if (value != _blkAttributes)
                {
                    _blkAttributes = value;
                    RaisePropertyChanged("blkAttributes");
                }
            }
        }
    }
}
