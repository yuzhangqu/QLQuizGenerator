using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportXlsToDownload
{
    public class Lv9 : Lv
    {
        public Lv9() : base(8)
        {
        }

        public override void Init(ref Random rand, bool mix)
        {
            if (mix)
            {
                OneDigit = 2;
                OneDigitMinus = 2;
                TwoDigit = 3;
                TwoDigitMinus = 1;

            }
            else
            {
                OneDigit = 4;
                TwoDigit = 4;
            }
            Generate(ref rand);
        }
    }
}