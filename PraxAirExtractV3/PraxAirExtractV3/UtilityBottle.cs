using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PraxAirExtractV3
{
    class UtilityBottle : Bottle
    { 
        public string AnalysisCylinder { get; set; }

        public string Concentration { get; set; }

        public UtilityBottle(string analysisCyl, string cylNum, string lot, string certDate, string concentration)
        {
            AnalysisCylinder = analysisCyl;

            CylinderNumber = cylNum;

            LotNumber = lot;

            CertificationDate = certDate;

            Concentration = concentration;


        }

        

        public override string ToString()
        {
            return AnalysisCylinder + " " + CylinderNumber + " " + LotNumber + " " + CertificationDate + " " + Concentration;
        }
    }

    
}
