using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportXlsToDownload
{
    public class Lv7 : Lv
    {
        public Lv7() : base(12)
        {
        }

        public override void Init(ref Random rand, bool mix)
        {
            if (mix)
            {
                OneDigit = 1;
                OneDigitMinus = 1;
                TwoDigit = 3;
                TwoDigitMinus = 1;
                ThreeDigit = 3;
                ThreeDigitMinus = 1;
                FourDigit = 1;
                FourDigitMinus = 1;
            }
            else
            {
                OneDigit = 2;
                TwoDigit = 4;
                ThreeDigit = 4;
                FourDigit = 2;
            }
            Generate(ref rand);
        }
    }
}