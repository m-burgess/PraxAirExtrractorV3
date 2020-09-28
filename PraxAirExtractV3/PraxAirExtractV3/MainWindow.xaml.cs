using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
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
        CalSheet calSheet = null;
        SpanBottle nox500 = null;
        SpanBottle nox2500 = null;
        SpanBottle nox10k = null;
        SpanBottle no500 = null;
        SpanBottle no2500 = null;
        SpanBottle no10k = null;
        SpanBottle thc500 = null;
        SpanBottle thc2500 = null;
        SpanBottle thc10k = null;
        SpanBottle ch4500 = null;
        SpanBottle ch42500 = null;
        SpanBottle ch410k = null;
        SpanBottle co5000 = null;
        SpanBottle coHigh = null;
        SpanBottle co2 = null;
        SpanBottle egr = null;
        SpanBottle o225 = null;
        UtilityBottle n21 = null;
        UtilityBottle n22 = null;
        UtilityBottle n23 = null;
        UtilityBottle n24 = null;
        UtilityBottle air1 = null;
        UtilityBottle air2 = null;
        UtilityBottle air3 = null;
        UtilityBottle air4 = null;
        UtilityBottle o21001 = null;
        UtilityBottle o21002 = null;
        UtilityBottle o21003 = null;
        UtilityBottle o21004 = null;
        UtilityBottle fuel1 = null;
        UtilityBottle fuel2 = null;

        


        public MainWindow()
        {
            InitializeComponent();

        }


        private void submitButton_Click(object sender, RoutedEventArgs e)
        {
            CalSheet calSheet = new CalSheet(calSheetTextBox.Text);


            proessBar.Visibility = Visibility.Visible;

            proessBar.Value = 5;

            //Nox
            if (nox500TextBox.Text != "")
            {
                PDF nox500Pdf = new PDF(directoryTextBox.Text + "\\" + nox500TextBox.Text);
                nox500 = nox500Pdf.nox500Extraction(nox500Pdf);
            }

            if (nox2500TextBox.Text != "")
            {
                PDF nox2500Pdf = new PDF(directoryTextBox.Text + "\\" + nox2500TextBox.Text);
                nox2500 = nox2500Pdf.nox2500Extraction(nox2500Pdf);
            }

            if (nox10kTextBox.Text != "")
            {
                PDF nox10kPdf = new PDF(directoryTextBox.Text + "\\" + nox10kTextBox.Text);
                nox10k = nox10kPdf.nox10kExtraction(nox10kPdf);
            }

            proessBar.Value = 10;

            //NO(C)
            if (no500TextBox.Text != "")
            {
                PDF no500Pdf = new PDF(directoryTextBox.Text + "\\" + no500TextBox.Text);
                no500 = no500Pdf.no500Extraction(no500Pdf);
            }

            if (no2500TextBox.Text != "")
            {
                PDF no2500Pdf = new PDF(directoryTextBox.Text + "\\" + no2500TextBox.Text);
                no2500 = no2500Pdf.no2500Extraction(no2500Pdf);
            }

            if (no10kTextBox.Text != "")
            {
                PDF no10kPdf = new PDF(directoryTextBox.Text + "\\" + no10kTextBox.Text);
                no10k = no10kPdf.no10kExtraction(no10kPdf);
            }

            proessBar.Value = 15;

            //THC
            if (thc500TextBox.Text != "")
            {
                PDF thc500Pdf = new PDF(directoryTextBox.Text + "\\" + thc500TextBox.Text);
                thc500 = thc500Pdf.thc500Extraction(thc500Pdf);
            }

            if (thc2500TextBox.Text != "")
            {
                PDF thc2500Pdf = new PDF(directoryTextBox.Text + "\\" + thc2500TextBox.Text);
                thc2500 = thc2500Pdf.thc2500Extraction(thc2500Pdf);
            }

            if (thc10kTextBox.Text != "")
            {
                PDF thc10kPdf = new PDF(directoryTextBox.Text + "\\" + thc10kTextBox.Text);
                thc10k = thc10kPdf.thc10kExtraction(thc10kPdf);
            }

            proessBar.Value = 20
                ;
            //CH4
            if (ch4500TextBox.Text != "")
            {
                PDF ch4500Pdf = new PDF(directoryTextBox.Text + "\\" + ch4500TextBox.Text);
                ch4500 = ch4500Pdf.ch4500Extraction(ch4500Pdf);
            }

            if (ch42500TextBox.Text != "")
            {
                PDF ch42500Pdf = new PDF(directoryTextBox.Text + "\\" + ch42500TextBox.Text);
                ch42500 = ch42500Pdf.ch42500Extraction(ch42500Pdf);
            }

            if (ch410kTextBox.Text != "")
            {
                PDF ch410kPdf = new PDF(directoryTextBox.Text + "\\" + ch410kTextBox.Text);
                ch410k = ch410kPdf.ch410kExtraction(ch410kPdf);
            }

            proessBar.Value = 25;

            //CO(L)
            if (co5000TextBox.Text != "")
            {
                PDF co5000Pdf = new PDF(directoryTextBox.Text + "\\" + co5000TextBox.Text);
                co5000 = co5000Pdf.co5000Extraction(co5000Pdf);
            }

            //CO(H)
            if (coHighTextBox.Text != "")
            {
                PDF coHighPdf = new PDF(directoryTextBox.Text + "\\" + coHighTextBox.Text);
                coHigh = coHighPdf.coHighExtraction(coHighPdf);
            }

            //CO2
            if (co2TextBox.Text != "")
            {
                PDF co2Pdf = new PDF(directoryTextBox.Text + "\\" + co2TextBox.Text);
                co2 = co2Pdf.co216Extraction(co2Pdf);
            }

            proessBar.Value = 30;

            //EGR
            if (egrTextBox.Text != "")
            {
                PDF egrPdf = new PDF(directoryTextBox.Text + "\\" + egrTextBox.Text);
                egr = egrPdf.egr16Extraction(egrPdf);
            }

            //O2 25%
            if (o225TextBox.Text != "")
            {
                PDF o225Pdf = new PDF(directoryTextBox.Text + "\\" + o225TextBox.Text);
                o225 = o225Pdf.o225Extraction(o225Pdf);
            }

            proessBar.Value = 35;

            //N2
            if (n21TextBox.Text != "")
            {
                string replacetext = n21TextBox.Text.Replace("N2 1","");
                PDF n21Pdf = new PDF(directoryTextBox.Text + "\\" + n21TextBox.Text);
                n21 = n21Pdf.n2Extraction(n21Pdf,replacetext.Replace(".pdf",""));
            }

            if (n22TextBox.Text != "")
            {
                string replacetext = n22TextBox.Text.Replace("N2 2", "");
                PDF n22Pdf = new PDF(directoryTextBox.Text + "\\" + n22TextBox.Text);
                n22 = n22Pdf.n2Extraction(n22Pdf,replacetext.Replace(".pdf", ""));
            }

            if (n23TextBox.Text != "")
            {
                string replacetext = n23TextBox.Text.Replace("N2 3", "");
                PDF n23Pdf = new PDF(directoryTextBox.Text + "\\" + n23TextBox.Text);
                n23 = n23Pdf.n2Extraction(n23Pdf, replacetext.Replace(".pdf", ""));
            }

            if (n24TextBox.Text != "")
            {
                string replacetext = n24TextBox.Text.Replace("N2 4", "");
                PDF n24Pdf = new PDF(directoryTextBox.Text + "\\" + n24TextBox.Text);
                n24 = n24Pdf.n2Extraction(n24Pdf,replacetext.Replace(".pdf", ""));
            }

            proessBar.Value = 40;

            //Air
            if (air1TextBox.Text != "")
            {
                string replacetext = air1TextBox.Text.Replace("Air 1", "");
                PDF air1Pdf = new PDF(directoryTextBox.Text + "\\" + air1TextBox.Text);
                air1 = air1Pdf.airExtraction(air1Pdf, replacetext.Replace(".pdf", ""));
            }

            if (air2TextBox.Text != "")
            {
                string replacetext = air2TextBox.Text.Replace("Air 2", "");
                PDF air2Pdf = new PDF(directoryTextBox.Text + "\\" + air2TextBox.Text);
                air2 = air2Pdf.airExtraction(air2Pdf, replacetext.Replace(".pdf", ""));
            }

            if (air3TextBox.Text != "")
            {
                string replacetext = air3TextBox.Text.Replace("Air 3", "");
                PDF air3Pdf = new PDF(directoryTextBox.Text + "\\" + air3TextBox.Text);
                air3 = air3Pdf.airExtraction(air3Pdf, replacetext.Replace(".pdf", ""));
            }

            if (air4TextBox.Text != "")
            {
                string replacetext = air4TextBox.Text.Replace("Air 4", "");
                PDF air4Pdf = new PDF(directoryTextBox.Text + "\\" + air4TextBox.Text);
                air4 = air4Pdf.airExtraction(air4Pdf, replacetext.Replace(".pdf", ""));
            }

            proessBar.Value = 45;

            //O2 100%
            if (o21001TextBox.Text != "")
            {
                string replacetext = o21001TextBox.Text.Replace("O2 100% 1", "");
                PDF o21001Pdf = new PDF(directoryTextBox.Text + "\\" + o21001TextBox.Text);
                o21001 = o21001Pdf.o2100Extraction(o21001Pdf, replacetext.Replace(".pdf", ""));
            }

            if (o21002TextBox.Text != "")
            {
                string replacetext = o21002TextBox.Text.Replace("O2 100% 2", "");
                PDF o21002Pdf = new PDF(directoryTextBox.Text + "\\" + o21002TextBox.Text);
                o21002 = o21002Pdf.o2100Extraction(o21002Pdf, replacetext.Replace(".pdf", ""));
            }

            if (o21003TextBox.Text != "")
            {
                string replacetext = o21003TextBox.Text.Replace("O2 100% 3", "");
                PDF o21003Pdf = new PDF(directoryTextBox.Text + "\\" + o21003TextBox.Text);
                o21003 = o21003Pdf.o2100Extraction(o21003Pdf, replacetext.Replace(".pdf", ""));
            }

            if (o21004TextBox.Text != "")
            {
                string replacetext = o21004TextBox.Text.Replace("O2 100% 4", "");
                PDF o21004Pdf = new PDF(directoryTextBox.Text + "\\" + o21004TextBox.Text);
                o21004 = o21004Pdf.o2100Extraction(o21004Pdf, replacetext.Replace(".pdf", ""));
            }

            proessBar.Value = 50;

            //Fuel
            if (fuel1TextBox.Text != "")
            {
                string replacetext = fuel1TextBox.Text.Replace("Fuel 1", "");
                PDF fuel1Pdf = new PDF(directoryTextBox.Text + "\\" + fuel1TextBox.Text);
                fuel1 = fuel1Pdf.fuelExtraction(fuel1Pdf, replacetext.Replace(".pdf", ""));
            }

            if (fuel2TextBox.Text != "")

            {
                string replacetext = fuel1TextBox.Text.Replace("Fuel 2", "");
                PDF fuel2Pdf = new PDF(directoryTextBox.Text + "\\" + fuel2TextBox.Text);
                fuel2 = fuel2Pdf.fuelExtraction(fuel2Pdf, replacetext.Replace(".pdf", ""));
            }

            proessBar.Value = 55;


            calSheet.WriteToCalSheet(calSheet.Path, nox500, nox2500, nox10k, no500, no2500,
                                     no10k, thc500, thc2500, thc10k, ch4500, ch42500, ch410k, co5000, coHigh, co2, egr, o225, n21, n22, n23, n24,
                                     air1, air2, air3, air4, o21001, o21002, o21003, o21004, fuel1, fuel2);
            proessBar.Value = 100;

            

            System.Windows.MessageBox.Show("Complete");
        }

        private void directoryButton_Click(object sender, RoutedEventArgs e)
        {

            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            // Show the FolderBrowserDialog.  
            DialogResult result = folderDlg.ShowDialog();
           
            directoryTextBox.Text = folderDlg.SelectedPath;
            Environment.SpecialFolder root = folderDlg.RootFolder;

            string[] filePaths = Directory.GetFiles(directoryTextBox.Text);

            foreach (string file in filePaths)
            {
                if (file.Contains("Nox 500"))
                {
                    nox500TextBox.Text = file.Replace(directoryTextBox.Text + "\\","");
                }

                if (file.Contains("Nox 2500"))
                {
                    nox2500TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("Nox 10k") || file.Contains("Nox 10000"))
                {
                    nox10kTextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("THC 500"))
                {
                    thc500TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("THC 2500"))
                {
                    thc2500TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("THC 10k") || file.Contains("THC 10000"))
                {
                    thc10kTextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("CH4 500"))
                {
                    ch4500TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("CH4 2500"))
                {
                    ch42500TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("CH4 10k") || file.Contains("CH4 10000"))
                {
                    ch410kTextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("CO(L) 5000"))
                {
                    co5000TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("CO(H) 1.6%"))
                {
                    coHighTextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("CO2 16%"))
                {
                    co2TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("EGR 16%"))
                {
                    egrTextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }


                if (file.Contains("O2 25%"))
                {
                    o225TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("N2 1"))
                {
                    n21TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("N2 2"))
                {
                    n22TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("N2 3"))
                {
                    n23TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("N2 4"))
                {
                    n24TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("Air 1"))
                {
                    air1TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("Air 2"))
                {
                    air2TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("Air 3"))
                {
                    air3TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("Air 4"))
                {
                    air4TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("O2 100% 1"))
                {
                    o21001TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("O2 100% 2"))
                {
                    o21002TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("O2 100% 3"))
                {
                    o21003TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("O2 100% 4"))
                {
                    o21004TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("Fuel 1"))
                {
                    fuel1TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }

                if (file.Contains("Fuel 2"))
                {
                    fuel2TextBox.Text = file.Replace(directoryTextBox.Text + "\\", "");
                }


            }

        }

        private void calSheetButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog folderDlg = new OpenFileDialog();
            // Show the FolderBrowserDialog.  
            DialogResult result = folderDlg.ShowDialog();

            calSheetTextBox.Text = folderDlg.FileName;
        }
    }
}
