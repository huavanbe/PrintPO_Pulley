using BFDataCrawler.View.ViewDetail;
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
    /// Interaction logic for MainWindow2.xaml
    /// </summary>
    public partial class MainWindow2 : Window
    {
        public MainWindow2()
        {
            InitializeComponent();

            MessageBoxResult dl = MessageBox.Show("Vui lòng chọn màn hình hiển thị cho ứng dụng:\n \n Yes: BW \n No: CNC","Message", MessageBoxButton.YesNo,MessageBoxImage.Question);
            if(dl==MessageBoxResult.Yes)
            {              
                BWWindow bw = new BWWindow();
                grid.Children.Add(bw);
            }
            else
            {
                CNCWindow cnc = new CNCWindow();
                grid.Children.Add(cnc);
            }
        }
    }
}
