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

namespace Guard_Inventory
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void viewInvButton_Click(object sender, RoutedEventArgs e)
        {
            mainLogo.Visibility = Visibility.Hidden;
            mainFrame.Content = new View_Inventory();
        }

        private void addRemoveButton_Click(object sender, RoutedEventArgs e)
        {
            mainLogo.Visibility = Visibility.Hidden;
            mainFrame.Content = new Add_Remove();
            
        }

        private void transLogButton_Click(object sender, RoutedEventArgs e)
        {
            mainLogo.Visibility = Visibility.Hidden;
            mainFrame.Content = new Trans_Log();
        }
    }
}
