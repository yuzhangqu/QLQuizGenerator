using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportXlsToDownload
{
    public class Lv10 : Lv
    {
        public Lv10() : base(6)
        {
        }

        public override void Init(ref Random rand, bool mix)
        {
            if (mix)
            {
                OneDigit = 2;
                OneDigitMinus = 2;
                TwoDigit = 2;

            }
            else
            {
                OneDigit = 4;
                TwoDigit = 2;
            }
            Generate(ref rand);
        }
    }
}