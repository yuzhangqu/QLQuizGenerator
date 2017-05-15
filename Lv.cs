using System;
using System.Collections.Generic;
using System.Linq;

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
        public int FiveDigit { get; set; }
        public int FiveDigitMinus { get; set; }

        private List<int> Ones = new List<int>();
        private List<int> Twos = new List<int>();
        private List<int> Threes = new List<int>();
        private List<int> Fours = new List<int>();
        private List<int> Fives = new List<int>();

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
            FiveDigit = 0;
            FiveDigitMinus = 0;

            Ones = new List<int>();
        }

        public abstract void Init(ref Random rand, bool mix = false);

        public void Generate(ref Random rand)
        {
            do
            {
                int index = 0;

                for (int i = 0; i < OneDigit; ++i)
                {
                    Nums[index++] = NextOne(ref rand);
                }

                for (int i = 0; i < OneDigitMinus; ++i)
                {
                    Nums[index++] = 0 - NextOne(ref rand);
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

                for (int i = 0; i < FiveDigit; ++i)
                {
                    Nums[index++] = rand.Next(10000, 100000);
                }

                for (int i = 0; i < FiveDigitMinus; ++i)
                {
                    Nums[index++] = 0 - rand.Next(10000, 100000);
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
            int old = Nums[0];
            int sum = old;
            if (sum < 0)
            {
                return true;
            }

            for (int i = 1; i < Count; ++i)
            {
                sum += Nums[i];
                if (sum < 0 || old == Nums[i] || old + Nums[i] == 0)
                {
                    return true;
                }
                old = Nums[i];
            }

            return false;
        }

        public int NextOne(ref Random rand)
        {
            if (Ones.Count == 0)
            {
                for (int i = 1; i < 10; ++i)
                {
                    Ones.Add(i);
                }

                for (int i = 8; i > 0; --i)
                {
                    int j = rand.Next(0, i + 1);
                    int temp = Ones[j];
                    Ones[j] = Ones[i];
                    Ones[i] = temp;
                }
            }

            int next = Ones[0];
            Ones.RemoveAt(0);

            return next;
        }

        public int NextTwo(ref Random rand)
        {
            if (Twos.Count == 0)
            {
                for (int i = 10; i < 100; ++i)
                {
                    Twos.Add(i);
                }

                for (int i = 89; i > 0; --i)
                {
                    int j = rand.Next(0, i + 1);
                    int temp = Twos[j];
                    Twos[j] = Twos[i];
                    Twos[i] = temp;
                }
            }

            int next = Twos[0];
            Twos.RemoveAt(0);

            return next;
        }

        public int NextThree(ref Random rand)
        {
            if (Threes.Count == 0)
            {
                for (int i = 100; i < 1000; ++i)
                {
                    Threes.Add(i);
                }

                for (int i = 899; i > 0; --i)
                {
                    int j = rand.Next(0, i + 1);
                    int temp = Threes[j];
                    Threes[j] = Threes[i];
                    Threes[i] = temp;
                }
            }

            int next = Threes[0];
            Threes.RemoveAt(0);

            return next;
        }

        public int NextFour(ref Random rand)
        {
            if (Fours.Count == 0)
            {
                for (int i = 1000; i < 10000; ++i)
                {
                    Fours.Add(i);
                }

                for (int i = 8999; i > 0; --i)
                {
                    int j = rand.Next(0, i + 1);
                    int temp = Fours[j];
                    Fours[j] = Fours[i];
                    Fours[i] = temp;
                }
            }

            int next = Fours[0];
            Fours.RemoveAt(0);

            return next;
        }

        public int NextFive(ref Random rand)
        {
            if (Fives.Count == 0)
            {
                for (int i = 10000; i < 100000; ++i)
                {
                    Fives.Add(i);
                }

                for (int i = 89999; i > 0; --i)
                {
                    int j = rand.Next(0, i + 1);
                    int temp = Fives[j];
                    Fives[j] = Fives[i];
                    Fives[i] = temp;
                }
            }

            int next = Fives[0];
            Fives.RemoveAt(0);

            return next;
        }
    }
}