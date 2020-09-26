using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PraxAirExtractV3
{
    class SpanBottle : Bottle
    {
       
        public string ExpirationDate { get; set; }

        public int Range { get; set; }

        public int SpanValue { get; set; }

        public SpanBottle(string gas, int range, string cylNum, string lot, string certDate, string expDate, int span)
        {
            CylinderNumber = cylNum;

            LotNumber = lot;

            CertificationDate = certDate;

            GasType = gas;

            ExpirationDate = expDate;

            Range = range;

            SpanValue = span;

        }

        public override string ToString()
        {
            return GasType + " " + Range + " " + CylinderNumber + " " + LotNumber + " " + CertificationDate + " " + ExpirationDate + " " + SpanValue;
        }



    }
}
