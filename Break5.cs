﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportXlsToDownload
{
    public class Break5 : Lv
    {
        public Break5() : base(5)
        {
        }

        public override void Init(ref Random rand, bool mix)
        {
            Nums[0] = rand.Next(1, 9);
            int result = Nums[0];
            for (int i = 1; i < Count; ++i)
            {
                if (result < 5)
                {
                    Nums[i] = rand.Next(5 - result, 5);
                    if (Nums[i] == Nums[i - 1] || Nums[i] + Nums[i - 1] == 0)
                    {
                        Nums[i] = rand.Next(1, 9 - result);
                    }
                }
                else
                {
                    Nums[i] = 0 - rand.Next(result - 4, 5);
                    if (Nums[i] == Nums[i - 1] || Nums[i] + Nums[i - 1] == 0)
                    {
                        Nums[i] = 0 - rand.Next(1, result);
                    }
                }

                result += Nums[i];
            }
        }
    }
}