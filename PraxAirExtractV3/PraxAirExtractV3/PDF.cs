using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;



namespace PraxAirExtractV3
{
    internal class PDF
    {
        public string Path { get; set; }

        public PDF(string path)
        {
            Path = path;
        }

        public string StripPDF(PDF filename)
        {
            //Load an existing PDF.

            PdfLoadedDocument loadedDocument = new PdfLoadedDocument(filename.Path);

            //Load the first page.

            PdfPageBase page = loadedDocument.Pages[0];

            //Extract text from first page.

            string extractedText = page.ExtractText();

            //Close the document

            loadedDocument.Close(true);

            return extractedText;
        }

        public SpanBottle nox500Extraction(PDF file)
        {
            SpanBottle nox500 = new SpanBottle(null, null, null, null, 0, "ppm", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains("For Reference Only:"))
                {
                    string removeNoxText = pdfList[i + 2].Replace("NOx ", "");
                    nox500.SpanValue = int.Parse(removeNoxText.Replace(" ppm", ""));
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    nox500.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    nox500.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    nox500.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    nox500.ExpirationDate = pdfList[i + 2];
                }
            }

            return nox500;
        }



        public SpanBottle nox2500Extraction(PDF file)
        {
            SpanBottle nox2500 = new SpanBottle(null, null, null, null, 0, "ppm", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains("For Reference Only:"))
                {
                    string removeNoxText = pdfList[i + 2].Replace("NOx ", "");
                    nox2500.SpanValue = int.Parse(removeNoxText.Replace(" ppm", ""));
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    nox2500.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    nox2500.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    nox2500.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    nox2500.ExpirationDate = pdfList[i + 2];
                }
            }

            return nox2500;
        }

        public SpanBottle nox10kExtraction(PDF file)
        {

            SpanBottle nox10k = new SpanBottle(null, null, null, null, 0, "ppm", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains("For Reference Only:"))
                {
                    string removeNoxText = pdfList[i + 2].Replace("NOx ", "");
                    nox10k.SpanValue = int.Parse(removeNoxText.Replace(" ppm", ""));
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    nox10k.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    nox10k.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    nox10k.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    nox10k.ExpirationDate = pdfList[i + 2];
                }
            }

            return nox10k;

        }

        public SpanBottle no500Extraction(PDF file)
        {

            SpanBottle no500 = new SpanBottle(null, null, null, null, 0, "ppm", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains("475 ppm"))
                {
                    string removePpm = pdfList[i + 2].Replace(" ppm", "");
                    no500.SpanValue = int.Parse(removePpm);
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    no500.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    no500.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    no500.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    no500.ExpirationDate = pdfList[i + 2];
                }
            }

            return no500;

        }

        public SpanBottle no2500Extraction(PDF file)
        {

            SpanBottle no2500 = new SpanBottle(null, null, null, null, 0, "ppm", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains(" 200 ppm"))
                {

                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    no2500.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    no2500.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    no2500.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    no2500.ExpirationDate = pdfList[i + 2];
                }
            }

            return no2500;

        }

        public SpanBottle no10kExtraction(PDF file)
        {

            SpanBottle no10k = new SpanBottle(null, null, null, null, 0, "ppm", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                if (pdfList[i].Contains("9500 ppm"))
                {
                    string removePpm = pdfList[i + 2].Replace(" ppm", "");
                    no10k.SpanValue = int.Parse(removePpm);
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    no10k.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    no10k.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    no10k.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    no10k.ExpirationDate = pdfList[i + 2];
                }
            }

            return no10k;

        }

        public SpanBottle thc500Extraction(PDF file)
        {

            SpanBottle thc500 = new SpanBottle(null, null, null, null, 0, "ppm", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains("158 ppm"))
                {
                    string removePpm = pdfList[i+2].Replace(" ppm", "");
                    double convertedInt = double.Parse(removePpm);
                    thc500.SpanValue = convertedInt*3;
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    thc500.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    thc500.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    thc500.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    thc500.ExpirationDate = pdfList[i + 2];
                }
            }

            return thc500;

        }

        public SpanBottle thc2500Extraction(PDF file)
        {

            SpanBottle thc2500 = new SpanBottle(null, null, null, null, 0, "ppm", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains(" 200 ppm"))
                {
                    thc2500.SpanValue = 0;
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    thc2500.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    thc2500.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    thc2500.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    thc2500.ExpirationDate = pdfList[i + 2];
                }
            }

            return thc2500;

        }

        public SpanBottle thc10kExtraction(PDF file)
        {

            SpanBottle thc10k = new SpanBottle(null, null, null, null, 0, "ppm", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains("3200 ppm"))
                {
                    string removePpm = pdfList[i+2].Replace(" ppm", "");
                    double convertedInt = double.Parse(removePpm);
                    thc10k.SpanValue = convertedInt * 3;
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    thc10k.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    thc10k.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    thc10k.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    thc10k.ExpirationDate = pdfList[i + 2];
                }
            }

            return thc10k;

        }

        public SpanBottle ch4500Extraction(PDF file)
        {

            SpanBottle ch4500 = new SpanBottle(null, null, null, null, 0, "ppm", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains(" 200 ppm"))
                {
                    ch4500.SpanValue = 0;
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    ch4500.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    ch4500.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    ch4500.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    ch4500.ExpirationDate = pdfList[i + 2];
                }
            }

            return ch4500;

        }

        public SpanBottle ch42500Extraction(PDF file)
        {

            SpanBottle ch42500 = new SpanBottle(null, null, null, null, 0, "ppm", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains(" 200 ppm"))
                {
                    ch42500.SpanValue = 0;
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    ch42500.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    ch42500.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    ch42500.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    ch42500.ExpirationDate = pdfList[i + 2];
                }
            }

            return ch42500;

        }

        public SpanBottle ch410kExtraction(PDF file)
        {

            SpanBottle ch410k = new SpanBottle(null, null, null, null, 0, "ppm", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains("9500 ppm"))
                {
                    string removePpm = pdfList[i + 2].Replace(" ppm","");
                    ch410k.SpanValue = double.Parse(removePpm);
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    ch410k.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    ch410k.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    ch410k.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    ch410k.ExpirationDate = pdfList[i + 2];
                }
            }

            return ch410k;

        }

        public SpanBottle co5000Extraction(PDF file)
        {

            SpanBottle co5000 = new SpanBottle(null, null, null, null, 0, "ppm", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains("4750 ppm"))
                {
                    string removePpm = pdfList[i + 2].Replace(" ppm","");
                    co5000.SpanValue = double.Parse(removePpm);
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    co5000.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    co5000.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    co5000.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    co5000.ExpirationDate = pdfList[i + 2];
                }
            }

            return co5000;

        }

        public SpanBottle coHighExtraction(PDF file)
        {

            SpanBottle coHigh = new SpanBottle(null, null, null, null, 0, "ppm", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains("1.52 %"))
                {
                    string removePercent = pdfList[i + 2].Replace(" %", "");
                    double coHighSpan = double.Parse(removePercent);
                    coHigh.SpanValue = coHighSpan * 10000;
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    coHigh.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    coHigh.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    coHigh.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    coHigh.ExpirationDate = pdfList[i + 2];
                }
            }

            return coHigh;

        }

        public SpanBottle co216Extraction(PDF file)
        {

            SpanBottle co216 = new SpanBottle(null, null, null, null, 0, "%", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains("15.2 %"))
                {
                    string removePercent = pdfList[i + 2].Replace(" %", "");
                    co216.SpanValue = double.Parse(removePercent);
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    co216.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    co216.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    co216.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    co216.ExpirationDate = pdfList[i + 2];
                }
            }

            return co216;

        }

        public SpanBottle egr16Extraction(PDF file)
        {
            SpanBottle egr16 = new SpanBottle(null, null, null, null, 0, "%", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains("15.2 %"))
                {
                    string removePercent = pdfList[i + 2].Replace(" %", "");
                    egr16.SpanValue = double.Parse(removePercent);
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    egr16.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    egr16.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    egr16.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    egr16.ExpirationDate = pdfList[i + 2];
                }
            }

            return egr16;
        }

        public SpanBottle o225Extraction(PDF file)
        {

            SpanBottle o225 = new SpanBottle(null, null, null, null, 0, "%", "1% NIST");
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains("21.0 %"))
                {
                    string removePercent = pdfList[1 + 2].Replace(" %","");
                    string removeSpace = removePercent.Replace(" ", "");
                    //o225.SpanValue = double.Parse(removeSpace);
                }

                //Cylinder Number
                if (pdfList[i].Contains("Cylinder Number(s):"))
                {

                    o225.CylinderNumber = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    o225.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    o225.CertificationDate = pdfList[i - 2];
                }

                //Expiration Date
                if (pdfList[i].Contains("Expiration Date:"))
                {
                    o225.ExpirationDate = pdfList[i + 2];
                }
            }

            
            return o225;

        }

        public UtilityBottle n2Extraction(PDF file, string cylinderNumber)
        {

            UtilityBottle n2 = new UtilityBottle(null,null,null,null,null);
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains("Analyzed Cylinder Number(s):"))
                {
                    n2.AnalysisCylinder = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    n2.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    n2.CertificationDate = pdfList[i - 2];
                }


            }

            //Cylinder Number
            n2.CylinderNumber = cylinderNumber;

            n2.Concentration = "";

            return n2;

        }

        public UtilityBottle airExtraction(PDF file, string cylinderNumber)
        {

            UtilityBottle air = new UtilityBottle(null, null, null, null, null);
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains("Analyzed Cylinder Number(s):"))
                {
                    air.AnalysisCylinder = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    air.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    air.CertificationDate = pdfList[i - 2];
                }

            }

            air.CylinderNumber = cylinderNumber;

            air.Concentration = "";

            return air;

        }

        public UtilityBottle o2100Extraction(PDF file, string cylinderNumber)
        {

            UtilityBottle o2100 = new UtilityBottle(null, null, null, null, null);
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains("Analyzed Cylinder Number(s):"))
                {
                    o2100.AnalysisCylinder = pdfList[i + 2];
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    o2100.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    o2100.CertificationDate = pdfList[i - 2];
                }

                //Concentration
                if (pdfList[i].Contains("99.993"))
                {
                    string removePercentage = pdfList[i + 2].Replace(" %", "");
                    o2100.Concentration = removePercentage;
                }

            }

            o2100.CylinderNumber = cylinderNumber;

            return o2100;

        }

        public UtilityBottle fuelExtraction(PDF file, string cylinderNumber)
        {

            UtilityBottle fuel = new UtilityBottle(null, null, null, null, null);
            //Extract Text From PDF
            String extractedText = StripPDF(file);
            //Split Text into List
            List<string> pdfList = new List<string>(extractedText.Split(new string[] { "\r\n" }, StringSplitOptions.None));

            //Loop through list and get specfic values
            for (int i = 0; i < pdfList.Count; i++)
            {
                //Span Value
                if (pdfList[i].Contains(" 200 ppm"))
                {
                    fuel.AnalysisCylinder = "";
                }

                //Lot Number
                if (pdfList[i].Contains("Lot Number:"))
                {
                    fuel.LotNumber = pdfList[i - 2];
                }

                //Certification Date
                if (pdfList[i].Contains("Certificate Issuance Date:"))
                {
                    fuel.CertificationDate = pdfList[i - 2];
                }

            }

            fuel.CylinderNumber = cylinderNumber;

            return fuel;

        }


    }
}

    

