using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace BlockAttributePrj
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class ModifyBlock : Window
    {
        public ModifyBlock()
        {
            InitializeComponent();
        }

        private void ListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
            ListView Listview = sender as ListView;
            if (Listview != null)
            {

                ModifyBlockViewModel vm = Listview.DataContext as ModifyBlockViewModel;
                if (vm!= null)
                {
                    vm.blkSelected = (BlockReferenceProperties)e.AddedItems[0];
                    vm.blkAttributes = vm.blkSelected.blkAttribute;
                }
            }

        }
    }
}
