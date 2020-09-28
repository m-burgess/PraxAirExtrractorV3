using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace PraxAirExtractV3
{
    class CalSheet
    {
        public string Path { get; set; }

        public CalSheet(string path)
        {
            Path = path;
        }
       
        public void WriteToCalSheet(string calSheetPath, SpanBottle nox500, SpanBottle nox2500, SpanBottle nox10k, SpanBottle no500,
                                          SpanBottle no2500, SpanBottle no10k, SpanBottle thc500, SpanBottle thc2500, SpanBottle thc10k,
                                          SpanBottle ch4500, SpanBottle ch42500, SpanBottle ch410k, SpanBottle co5000, SpanBottle coHigh,
                                          SpanBottle co216, SpanBottle egr16, SpanBottle o225, UtilityBottle n21, UtilityBottle n22,
                                          UtilityBottle n23, UtilityBottle n24, UtilityBottle air1, UtilityBottle air2, UtilityBottle air3,
                                          UtilityBottle air4, UtilityBottle o21001, UtilityBottle o21002, UtilityBottle o21003, UtilityBottle o21004,
                                          UtilityBottle fuel1, UtilityBottle fuel2)
        {

            //Create Application, Workbook, and Worksheet
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(calSheetPath);
            Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets.get_Item(2);
            //Write values in cell locations
            
            //Nox 500
            excelWorksheet.Cells[5, 2] = nox500.CylinderNumber;
            excelWorksheet.Cells[5, 3] = nox500.LotNumber;
            excelWorksheet.Cells[5, 4] = nox500.CertificationDate;
            excelWorksheet.Cells[5, 5] = nox500.ExpirationDate;
            excelWorksheet.Cells[5, 6] = nox500.SpanValue;
            excelWorksheet.Cells[5, 7] = nox500.Unit;
            excelWorksheet.Cells[5, 8] = nox500.Tracability;
            //Cylinder Number


            //Nox 2500
            excelWorksheet.Cells[6, 2] = nox2500?.CylinderNumber;
            excelWorksheet.Cells[6, 3] = nox2500?.LotNumber;
            excelWorksheet.Cells[6, 4] = nox2500?.CertificationDate;
            excelWorksheet.Cells[6, 5] = nox2500?.ExpirationDate;
            excelWorksheet.Cells[6, 6] = nox2500?.SpanValue;
            excelWorksheet.Cells[6, 7] = nox2500?.Unit;
            excelWorksheet.Cells[6, 8] = nox2500?.Tracability;

            //Nox 10k
            excelWorksheet.Cells[7, 2] = nox10k?.CylinderNumber;
            excelWorksheet.Cells[7, 3] = nox10k?.LotNumber;
            excelWorksheet.Cells[7, 4] = nox10k?.CertificationDate;
            excelWorksheet.Cells[7, 5] = nox10k?.ExpirationDate;
            excelWorksheet.Cells[7, 6] = nox10k?.SpanValue;
            excelWorksheet.Cells[7, 7] = nox10k?.Unit;
            excelWorksheet.Cells[7, 8] = nox10k?.Tracability;

            //No 500
            excelWorksheet.Cells[8, 2] = no500?.CylinderNumber;
            excelWorksheet.Cells[8, 3] = no500?.LotNumber;
            excelWorksheet.Cells[8, 4] = no500?.CertificationDate;
            excelWorksheet.Cells[8, 5] = no500?.ExpirationDate;
            excelWorksheet.Cells[8, 6] = no500?.SpanValue;
            excelWorksheet.Cells[8, 7] = no500?.Unit;
            excelWorksheet.Cells[8, 8] = no500?.Tracability;
            //Cylinder Number


            //No 2500
            excelWorksheet.Cells[9, 2] = no2500?.CylinderNumber;
            excelWorksheet.Cells[9, 3] = no2500?.LotNumber;
            excelWorksheet.Cells[9, 4] = no2500?.CertificationDate;
            excelWorksheet.Cells[9, 5] = no2500?.ExpirationDate;
            excelWorksheet.Cells[9, 6] = no2500?.SpanValue;
            excelWorksheet.Cells[9, 7] = no2500?.Unit;
            excelWorksheet.Cells[9, 8] = no2500?.Tracability;

            //No 10k
            excelWorksheet.Cells[10, 2] = no10k?.CylinderNumber;
            excelWorksheet.Cells[10, 3] = no10k?.LotNumber;
            excelWorksheet.Cells[10, 4] = no10k?.CertificationDate;
            excelWorksheet.Cells[10, 5] = no10k?.ExpirationDate;
            excelWorksheet.Cells[10, 6] = no10k?.SpanValue;
            excelWorksheet.Cells[10, 7] = no10k?.Unit;
            excelWorksheet.Cells[10, 8] = no10k?.Tracability;

            //THC 500
            excelWorksheet.Cells[11, 2] = thc500?.CylinderNumber;
            excelWorksheet.Cells[11, 3] = thc500?.LotNumber;
            excelWorksheet.Cells[11, 4] = thc500?.CertificationDate;
            excelWorksheet.Cells[11, 5] = thc500?.ExpirationDate;
            excelWorksheet.Cells[11, 6] = thc500?.SpanValue;
            excelWorksheet.Cells[11, 7] = thc500?.Unit;
            excelWorksheet.Cells[11, 8] = thc500?.Tracability;
            //Cylinder Number


            //THC 2500
            excelWorksheet.Cells[12, 2] = thc2500?.CylinderNumber;
            excelWorksheet.Cells[12, 3] = thc2500?.LotNumber;
            excelWorksheet.Cells[12, 4] = thc2500?.CertificationDate;
            excelWorksheet.Cells[12, 5] = thc2500?.ExpirationDate;
            excelWorksheet.Cells[12, 6] = thc2500?.SpanValue;
            excelWorksheet.Cells[12, 7] = thc2500?.Unit;
            excelWorksheet.Cells[12, 8] = thc2500?.Tracability;

            //THC 10k
            excelWorksheet.Cells[13, 2] = thc10k?.CylinderNumber;
            excelWorksheet.Cells[13, 3] = thc10k?.LotNumber;
            excelWorksheet.Cells[13, 4] = thc10k?.CertificationDate;
            excelWorksheet.Cells[13, 5] = thc10k?.ExpirationDate;
            excelWorksheet.Cells[13, 6] = thc10k?.SpanValue;
            excelWorksheet.Cells[13, 7] = thc10k?.Unit;
            excelWorksheet.Cells[13, 8] = thc10k?.Tracability;

            //Ch4 500
            excelWorksheet.Cells[14, 2] = ch4500?.CylinderNumber;
            excelWorksheet.Cells[14, 3] = ch4500?.LotNumber;
            excelWorksheet.Cells[14, 4] = ch4500?.CertificationDate;
            excelWorksheet.Cells[14, 5] = ch4500?.ExpirationDate;
            excelWorksheet.Cells[14, 6] = ch4500?.SpanValue;
            excelWorksheet.Cells[14, 7] = ch4500?.Unit;
            excelWorksheet.Cells[14, 8] = ch4500?.Tracability;
            //Cylinder Number


            //CH4 2500
            excelWorksheet.Cells[15, 2] = ch42500?.CylinderNumber;
            excelWorksheet.Cells[15, 3] = ch42500?.LotNumber;
            excelWorksheet.Cells[15, 4] = ch42500?.CertificationDate;
            excelWorksheet.Cells[15, 5] = ch42500?.ExpirationDate;
            excelWorksheet.Cells[15, 6] = ch42500?.SpanValue;
            excelWorksheet.Cells[15, 7] = ch42500?.Unit;
            excelWorksheet.Cells[15, 8] = ch42500?.Tracability;

            //CH4 10k
            excelWorksheet.Cells[16, 2] = ch410k?.CylinderNumber;
            excelWorksheet.Cells[16, 3] = ch410k?.LotNumber;
            excelWorksheet.Cells[16, 4] = ch410k?.CertificationDate;
            excelWorksheet.Cells[16, 5] = ch410k?.ExpirationDate;
            excelWorksheet.Cells[16, 6] = ch410k?.SpanValue;
            excelWorksheet.Cells[16, 7] = ch410k?.Unit;
            excelWorksheet.Cells[16, 8] = ch410k?.Tracability;

            //CO(L)
            excelWorksheet.Cells[17, 2] = co5000?.CylinderNumber;
            excelWorksheet.Cells[17, 3] = co5000?.LotNumber;
            excelWorksheet.Cells[17, 4] = co5000?.CertificationDate;
            excelWorksheet.Cells[17, 5] = co5000?.ExpirationDate;
            excelWorksheet.Cells[17, 6] = co5000?.SpanValue;
            excelWorksheet.Cells[17, 7] = co5000?.Unit;
            excelWorksheet.Cells[17, 8] = co5000?.Tracability;
            //Cylinder Number


            //CO(H)
            excelWorksheet.Cells[18, 2] = coHigh?.CylinderNumber;
            excelWorksheet.Cells[18, 3] = coHigh?.LotNumber;
            excelWorksheet.Cells[18, 4] = coHigh?.CertificationDate;
            excelWorksheet.Cells[18, 5] = coHigh?.ExpirationDate;
            excelWorksheet.Cells[18, 6] = coHigh?.SpanValue;
            excelWorksheet.Cells[18, 7] = coHigh?.Unit;
            excelWorksheet.Cells[18, 8] = coHigh?.Tracability;

            //CO2
            excelWorksheet.Cells[19, 2] = co216?.CylinderNumber;
            excelWorksheet.Cells[19, 3] = co216?.LotNumber;
            excelWorksheet.Cells[19, 4] = co216?.CertificationDate;
            excelWorksheet.Cells[19, 5] = co216?.ExpirationDate;
            excelWorksheet.Cells[19, 6] = co216?.SpanValue;
            excelWorksheet.Cells[19, 7] = co216?.Unit;
            excelWorksheet.Cells[19, 8] = co216?.Tracability;

            //EGR
            excelWorksheet.Cells[20, 2] = egr16?.CylinderNumber;
            excelWorksheet.Cells[20, 3] = egr16?.LotNumber;
            excelWorksheet.Cells[20, 4] = egr16?.CertificationDate;
            excelWorksheet.Cells[20, 5] = egr16?.ExpirationDate;
            excelWorksheet.Cells[20, 6] = egr16?.SpanValue;
            excelWorksheet.Cells[20, 7] = egr16?.Unit;
            excelWorksheet.Cells[20, 8] = egr16?.Tracability;

            //O2 25%
            excelWorksheet.Cells[21, 2] = o225?.CylinderNumber;
            excelWorksheet.Cells[21, 3] = o225?.LotNumber;
            excelWorksheet.Cells[21, 4] = o225?.CertificationDate;
            excelWorksheet.Cells[21, 5] = o225?.ExpirationDate;
            excelWorksheet.Cells[21, 6] = o225?.SpanValue;
            excelWorksheet.Cells[21, 7] = o225?.Unit;
            excelWorksheet.Cells[21, 8] = o225?.Tracability;

            //N2 1
            excelWorksheet.Cells[23, 4] = n21?.AnalysisCylinder;
            excelWorksheet.Cells[24, 4] = n21?.CylinderNumber;
            excelWorksheet.Cells[25, 4] = n21?.LotNumber;
            excelWorksheet.Cells[26, 4] = n21?.CertificationDate;

            //N2 2
            excelWorksheet.Cells[23, 5] = n22?.AnalysisCylinder;
            excelWorksheet.Cells[24, 5] = n22?.CylinderNumber;
            excelWorksheet.Cells[25, 5] = n22?.LotNumber;
            excelWorksheet.Cells[26, 5] = n22?.CertificationDate;

            //N2 3
            excelWorksheet.Cells[23, 6] = n23?.AnalysisCylinder;
            excelWorksheet.Cells[24, 6] = n23?.CylinderNumber;
            excelWorksheet.Cells[25, 6] = n23?.LotNumber;
            excelWorksheet.Cells[26, 6] = n23?.CertificationDate;

            //N2 4
            excelWorksheet.Cells[23, 7] = n24?.AnalysisCylinder;
            excelWorksheet.Cells[24, 7] = n24?.CylinderNumber;
            excelWorksheet.Cells[25, 7] = n24?.LotNumber;
            excelWorksheet.Cells[26, 7] = n24?.CertificationDate;

            //Air 1
            excelWorksheet.Cells[29, 4] = air1?.AnalysisCylinder;
            excelWorksheet.Cells[30, 4] = air1?.CylinderNumber;
            excelWorksheet.Cells[31, 4] = air1?.LotNumber;
            excelWorksheet.Cells[32, 4] = air1?.CertificationDate;

            //Air 2
            excelWorksheet.Cells[29, 5] = air2?.AnalysisCylinder;
            excelWorksheet.Cells[30, 5] = air2?.CylinderNumber;
            excelWorksheet.Cells[31, 5] = air2?.LotNumber;
            excelWorksheet.Cells[32, 5] = air2?.CertificationDate;

            //Air 3
            excelWorksheet.Cells[29, 6] = air3?.AnalysisCylinder;
            excelWorksheet.Cells[30, 6] = air3?.CylinderNumber;
            excelWorksheet.Cells[31, 6] = air3?.LotNumber;
            excelWorksheet.Cells[32, 6] = air3?.CertificationDate;


            //Air 4
            excelWorksheet.Cells[29, 7] = air4?.AnalysisCylinder;
            excelWorksheet.Cells[30, 7] = air4?.CylinderNumber;
            excelWorksheet.Cells[31, 7] = air4?.LotNumber;
            excelWorksheet.Cells[32, 7] = air4?.CertificationDate;

            //O2 100% 1
            excelWorksheet.Cells[41, 4] = o21001?.AnalysisCylinder;
            excelWorksheet.Cells[42, 4] = o21001?.CylinderNumber;
            excelWorksheet.Cells[43, 4] = o21001?.LotNumber;
            excelWorksheet.Cells[44, 4] = o21001?.CertificationDate;
            excelWorksheet.Cells[46, 4] = o21001?.Concentration;

            //O2 100% 2
            excelWorksheet.Cells[41, 5] = o21002?.AnalysisCylinder;
            excelWorksheet.Cells[42, 5] = o21002?.CylinderNumber;
            excelWorksheet.Cells[43, 5] = o21002?.LotNumber;
            excelWorksheet.Cells[44, 5] = o21002?.CertificationDate;
            excelWorksheet.Cells[46, 5] = o21002?.Concentration;

            //O2 100% 3
            excelWorksheet.Cells[41, 6] = o21003?.AnalysisCylinder;
            excelWorksheet.Cells[42, 6] = o21003?.CylinderNumber;
            excelWorksheet.Cells[43, 6] = o21003?.LotNumber;
            excelWorksheet.Cells[44, 6] = o21003?.CertificationDate;
            excelWorksheet.Cells[46, 6] = o21003?.Concentration;

            //O2 100% 4
            excelWorksheet.Cells[41, 7] = o21004?.AnalysisCylinder;
            excelWorksheet.Cells[42, 7] = o21004?.CylinderNumber;
            excelWorksheet.Cells[43, 7] = o21004?.LotNumber;
            excelWorksheet.Cells[44, 7] = o21004?.CertificationDate;
            excelWorksheet.Cells[46, 7] = o21004?.Concentration;

            //Fuel 1
            excelWorksheet.Cells[35, 4] = fuel1?.AnalysisCylinder;
            excelWorksheet.Cells[36, 4] = fuel1?.CylinderNumber;
            excelWorksheet.Cells[37, 4] = fuel1?.LotNumber;
            excelWorksheet.Cells[38, 4] = fuel1?.CertificationDate;
            //Fuel 2
            excelWorksheet.Cells[35, 5] = fuel2?.AnalysisCylinder;
            excelWorksheet.Cells[36, 5] = fuel2?.CylinderNumber;
            excelWorksheet.Cells[37, 5] = fuel2?.LotNumber;
            excelWorksheet.Cells[38, 5] = fuel2?.CertificationDate;



            //Save, Close, and Quit Excel
            excelWorkbook.Save();
            excelWorkbook.Close();
            excelApp.Quit();

        }

       
    }
}
