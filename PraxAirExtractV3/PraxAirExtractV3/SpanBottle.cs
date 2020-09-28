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

        public double SpanValue { get; set; }

        public SpanBottle(string cylNum, string lot, string certDate, string expDate, double span, string unit, string traceability)
        {
            CylinderNumber = cylNum;

            LotNumber = lot;

            CertificationDate = certDate;

            Unit = unit;

            ExpirationDate = expDate;

            SpanValue = span;

            Tracability = traceability;

        }

        public override string ToString()
        {
            return CylinderNumber + " " + LotNumber + " " + CertificationDate + " " + ExpirationDate + " " + SpanValue + " " + Unit + " " + Tracability;
        }



    }
}
