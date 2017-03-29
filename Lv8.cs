using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportXlsToDownload
{
    public class Lv8 : Lv
    {
        public Lv8() : base(10)
        {
        }

        public override void Init(ref Random rand, bool mix)
        {
            if (mix)
            {
                OneDigit = 1;
                OneDigitMinus = 2;
                TwoDigit = 3;
                TwoDigitMinus = 1;
                ThreeDigit = 2;
                ThreeDigitMinus = 1;
            }
            else
            {
                OneDigit = 3;
                TwoDigit = 4;
                ThreeDigit = 3;
            }
            Generate(ref rand);
        }
    }
}