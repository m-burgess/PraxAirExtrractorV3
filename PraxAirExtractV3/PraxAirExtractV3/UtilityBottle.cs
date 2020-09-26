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

        int BottleNumber { get; set; }

        public UtilityBottle(string gas, int number, string cylNum, string lot, string certDate)
        {
            CylinderNumber = cylNum;

            LotNumber = lot;

            CertificationDate = certDate;

            GasType = gas;

            BottleNumber = number;


        }

        public override string ToString()
        {
            return GasType + " " + BottleNumber + " " + CylinderNumber + " " + LotNumber + " " + CertificationDate;
        }
    }

    
}
