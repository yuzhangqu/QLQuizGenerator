using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportXlsToDownload
{
    public abstract class Lv
    {
        public int Count { get; }
        public int[] Nums { get; }
        public int OneDigit { get; set; }
        public int OneDigitMinus { get; set; }
        public int TwoDigit { get; set; }
        public int TwoDigitMinus { get; set; }
        public int ThreeDigit { get; set; }
        public int ThreeDigitMinus { get; set; }
        public int FourDigit { get; set; }
        public int FourDigitMinus { get; set; }

        public Lv(int count)
        {
            Count = count;
            Nums = new int[Count];
            OneDigit = 0;
            OneDigitMinus = 0;
            TwoDigit = 0;
            TwoDigitMinus = 0;
            ThreeDigit = 0;
            ThreeDigitMinus = 0;
            FourDigit = 0;
            FourDigitMinus = 0;
        }

        public abstract void Init(ref Random rand, bool mix = false);

        public void Generate(ref Random rand)
        {
            do
            {
                int index = 0;

                for (int i = 0; i < OneDigit; ++i)
                {
                    Nums[index++] = rand.Next(1, 10);
                }

                for (int i = 0; i < OneDigitMinus; ++i)
                {
                    Nums[index++] = 0 - rand.Next(1, 10);
                }

                for (int i = 0; i < TwoDigit; ++i)
                {
                    Nums[index++] = rand.Next(10, 100);
                }

                for (int i = 0; i < TwoDigitMinus; ++i)
                {
                    Nums[index++] = 0 - rand.Next(10, 100);
                }

                for (int i = 0; i < ThreeDigit; ++i)
                {
                    Nums[index++] = rand.Next(100, 1000);
                }

                for (int i = 0; i < ThreeDigitMinus; ++i)
                {
                    Nums[index++] = 0 - rand.Next(100, 1000);
                }

                for (int i = 0; i < FourDigit; ++i)
                {
                    Nums[index++] = rand.Next(1000, 10000);
                }

                for (int i = 0; i < FourDigitMinus; ++i)
                {
                    Nums[index++] = 0 - rand.Next(1000, 10000);
                }
            } while (Negative());

            // Shuffle until valid
            do
            {
                for (int i = Count - 1; i > 0; --i)
                {
                    int j = rand.Next(0, i + 1);
                    Swap(ref Nums[j], ref Nums[i]);
                }
            } while (Invalid());
        }

        public void Swap(ref int i, ref int j)
        {
            int temp = i;
            i = j;
            j = temp;
        }

        public bool Negative()
        {
            int sum = 0;
            foreach (int i in Nums)
            {
                sum += i;
            }

            return sum < 0;
        }

        public bool Invalid()
        {
            int sum = 0;
            foreach (int i in Nums)
            {
                sum += i;
                if (sum < 0)
                {
                    return true;
                }
            }

            return false;
        }
    }
}