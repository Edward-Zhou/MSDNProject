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

namespace ExcelAddIn
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }
        private void cmDuplicateWB_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Duplicate");
        }

        private void cmRenameWB_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Rename");

        }

        private void cmShowFolder_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Show in Folder");
        }

        private void MyButton_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.MyButton.Focus();
        }

        private void MyButton_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
