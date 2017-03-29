using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportXlsToDownload
{
    public class Lv6 : Lv
    {
        public Lv6() : base(14)
        {
        }

        public override void Init(ref Random rand, bool mix)
        {
            if (mix)
            {
                TwoDigit = 3;
                TwoDigitMinus = 3;
                ThreeDigit = 3;
                ThreeDigitMinus = 1;
                FourDigit = 3;
                FourDigitMinus = 1;
            }
            else
            {
                TwoDigit = 6;
                ThreeDigit = 4;
                FourDigit = 4;
            }
            Generate(ref rand);
        }
    }
}