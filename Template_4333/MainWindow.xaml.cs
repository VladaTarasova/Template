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

namespace Template_4333
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var form = new _4333_Gibadulllina();
            form.Show();
            this.Hide();
        }

        private void Amir_Click(object sender, RoutedEventArgs e)
        {
            var form = new _4333_Gallyamov();
            form.Show();
            this.Hide();
        }

        private void Dina_Click(object sender, RoutedEventArgs e)
        {
            _4333_Sagitova sag = new _4333_Sagitova();
            sag.Show();
        }
        private void _4333_Davliev(object sender, RoutedEventArgs e)
        {
            _4333_Davliev davliev = new _4333_Davliev();
            davliev.Show();
        }

        private void dinarClick(object sender, RoutedEventArgs e)
        {
            var dinar = new _4333_Valiakhmetov();
            dinar.Show();
            this.Hide();
        }

        private void _4333_Ibragimov(object sender, RoutedEventArgs e)
        {
            _4333_Ibragimov ibragimov = new _4333_Ibragimov();
            ibragimov.Show();
            this.Hide();
        }
        private void _4333_Tarasova(object sender, RoutedEventArgs e)
        {
            _4333_Tarasova tarasova = new _4333_Tarasova();
            tarasova.Show();
        }
    }
}
