using CanMakeTree.DataBase;
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

namespace CanMakeTree
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly ChemicalContext _context = new ChemicalContext();

        public MainWindow()
        {
            InitializeComponent();
        }


        //按钮点击
        private void button_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("1111");
        }
    }
}
