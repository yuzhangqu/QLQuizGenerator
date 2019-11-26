using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportXlsToDownload
{
    public class XX7 : Lv
    {
        public XX7() : base(10)
        {
            Scale = 0.01;
        }

        public override void Init(ref Random rand, bool mix)
        {
            if (mix)
            {
                TwoDigit = 2;
                TwoDigitMinus = 1;
                ThreeDigit = 3;
                ThreeDigitMinus = 1;
                FourDigit = 2;
                FourDigitMinus = 1;
            }
            else
            {
                TwoDigit = 3;
                ThreeDigit = 4;
                FourDigit = 3;
            }
            Generate(ref rand);
        }
    }
}