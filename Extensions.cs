using System;
using System.Collections.Generic;


namespace ZZ11ToLetter
{
    public static class Extensions
    {
        public static IEnumerable<string> Split(this string str, int n)
        {
            if (String.IsNullOrEmpty(str) || n < 1)
                throw new ArgumentException();

            for (int i = 0; i < str.Length; i += n)
                yield return str.Substring(i, Math.Min(n, str.Length - i));
        }
    }
}
