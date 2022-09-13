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
using System.Windows.Shapes;

namespace BFDataCrawler.View
{
    /// <summary>
    /// Interaction logic for UnlockWindow.xaml
    /// </summary>
    public partial class UnlockWindow : Window
    {
        public UnlockWindow()
        {
            InitializeComponent();

        }

        private void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            
        }

        private void PasswordBox_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                if (this.DataContext != null)
                {
                    ((dynamic)this.DataContext).PasswordUnlock = ((PasswordBox)sender).Password;
                }
            }
            
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            txtPass.Focus();
        }
    }
}
