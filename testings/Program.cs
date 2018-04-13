using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace testings
{
    public class Program
    {
        static void Main(string[] args)
        {
            string inputString = "Предмет; 00.00.00г.; -хуй; Предмет; 00.00.00г.; -пизда;";
            string result = ReplaceSeparators(inputString, '-', ';', '|');

            Console.WriteLine(result);
            Console.ReadKey();
        }

        public static string ReplaceSeparators(string input, char startingChar, char oldChar, char newChar)
        {
            List<int> indexes = new List<int>();
            char[] chArr = input.ToCharArray();
            for (int i = 0; i < chArr.Length; i++)
            {
                if (chArr[i] == '-')
                {
                    int index = i;
                    while (chArr[i] != ';')
                        i++;
                    index = i;
                    indexes.Add(i);
                }
            }

            for (int i = 0; i < chArr.Length; i++)
            {
                if (indexes.Contains(i)) chArr[i] = '|';
            }
            string s = new string(chArr);
            return s;
        }
    }
}
