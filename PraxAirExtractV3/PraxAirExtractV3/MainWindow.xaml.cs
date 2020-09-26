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

namespace PraxAirExtractV3
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            SpanBottle nox1 = new SpanBottle("Nox", 500, "CC124356", "70089645", "12/8/20", "12/7/24", 478);

            Console.WriteLine(nox1.ToString());

            UtilityBottle n21 = new UtilityBottle("N2", 1, "1684009", "7000465", "3/5/21");

            Console.WriteLine(n21.LotNumber);
        }
    }
}
