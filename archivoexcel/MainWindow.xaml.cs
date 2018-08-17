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
namespace archivoexcel
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void CrearExcel(object sender, RoutedEventArgs e)
        {
            List<persona> personas = DatosDummy();
            ManejoExcel excel = new ManejoExcel(personas);
        }

        private List<persona> DatosDummy()
        {
            List<persona> lista = new List<persona>();

            persona jaime = new persona("123","jaime","gomez", "pasto","7312234","12/02/1988");
            lista.Add(jaime);
            persona juan = new persona("124", "juan", "torrez", "bogota", "7315634", "10/01/1987");
            lista.Add(juan);
            return lista;
            

        }
    }
}
