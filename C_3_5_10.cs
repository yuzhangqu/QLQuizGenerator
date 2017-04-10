using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportXlsToDownload
{
    public class C_3_5_10 : Lv
    {
        public C_3_5_10() : base(10)
        {
        }

        public override void Init(ref Random rand, bool mix)
        {
            if (mix)
            {
                ThreeDigit = 2;
                ThreeDigitMinus = 1;
                FourDigit = 2;
                FourDigitMinus = 1;
                FiveDigit = 3;
                FiveDigitMinus = 1;
            }
            else
            {
                ThreeDigit = 3;
                FourDigit = 3;
                FiveDigit = 4;
            }
            Generate(ref rand);
        }
    }
}