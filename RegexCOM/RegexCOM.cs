using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

namespace RegexCOM
{

    public enum RegexOptionsValues
    {
        /*
         * Enumeration of the different RegexOptions items. When calling this from VBA,
         * you can combine options by adding them together
        */

        None = 0,
        IgnoreCase = 1,
        Multiline = 2,
        ExplicitCapture = 4,
        Compiled = 8,
        Singleline = 16,
        IgnorePatternWhitespace = 32,
        RightToLeft = 64,
        ECMAScript = 256,
        CultureInvariant = 512,
        NonBacktracking = 1024
    }

    [Guid("C3AA9ABB-7A45-42B6-8729-CD539612DC2A")]
    public interface IRegX
    {

        bool IsMatch(string inputText, string pattern, RegexOptionsValues options = 0);

        string Match(string inputText, string pattern, RegexOptionsValues options = 0);

        string[] Matches(string inputText, string pattern, RegexOptionsValues options = 0);

        string[,] MatchesWithGroups(string inputText, string pattern, RegexOptionsValues options = 0);

        string Replace(string inputText, string pattern, string replacement, RegexOptionsValues options = 0);

        string[] Split(string inputText, string pattern, RegexOptionsValues options = 0);
    }

    [ComVisible(true)]
    [Guid("F8AF9ED7-2A6D-4DE8-8C5F-3DF8FBDE0921")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("RegexCOM.RegX")]
    public class RegX : IRegX
    {

        public bool IsMatch(string inputText, string pattern, RegexOptionsValues options = 0)
        {
            /*
             * Tests whether there ius at least one match in the input string, and returns
             * boolean true/false
            */

            return Regex.IsMatch(inputText, $@"{pattern}", (RegexOptions)options);
        }

        public string Match(string inputText, string pattern, RegexOptionsValues options = 0)
        {
            /*
             * Applies a pattern to the input string, and returns the first match if applicable,
             * or an empty string if there is no match
            */

            return Regex.Match(inputText, $@"{pattern}", (RegexOptions)options).ToString();
        }

        public string[] Matches(string inputText, string pattern, RegexOptionsValues options = 0)
        {
            /*
             * Applies a pattern to the input string, and if there is at least one match, 
             * returns a zero-based array of the matches. If there is no match, returns 
             * an empty array (VBA will treat this as an array with a UBound of -1)
            */

            MatchCollection m = Regex.Matches(inputText, $@"{pattern}", (RegexOptions)options);
            string[] result = new string[m.Count];
            for (int i = 0; i < m.Count; i++)
            {
                result[i] = m[i].ToString();
            }
            return result;
        }

        public string[,] MatchesWithGroups(string inputText, string pattern, RegexOptionsValues options = 0)
        {
            /*
             * Applies a pattern to the input string, and if there is at least one match, 
             * returns a multi-dimensional array.
             * 1) If there are no match groups, for n matches:
             *      array will have two dimensions, first upper bound (# matches - 1) and 
             *      second upper bound 0. array(n - 1, 0) will return the nth match as a whole
             * 2) If there are one or more match groups, for n matches:
             *      array will have two dimensions, first upper bound (# matches - 1) and
             *      second upper bound (# match groups). array(n - 1, 0) will return the
             *      match as a whole. array(n - 1, p) will return the value of the pth match
             *      group of the nth match
             * 
             * if there are no matches, returns an empty two-dimensional array (VBA will treat
             * each dimension's UBound as -1)
            */

            string[,] result = new string[0, 0];
            
            MatchCollection mc = Regex.Matches(inputText, $@"{pattern}", (RegexOptions)options);
            if (mc.Count > 0)
            {
                int numGroups = mc[0].Groups.Count;
                result = new string[mc.Count, numGroups];
                for (int i = 0; i < mc.Count; i++) 
                {
                    for (int j = 0; j < numGroups; j++)
                    {
                        result[i, j] = mc[i].Groups[j].Value;
                    }
                }
            }
            
            return result;
        }

        public string Replace(string inputText, string pattern, string replacement, RegexOptionsValues options = 0)
        {
            /*
             * Applies a pattern to the input string, and for each match, replaces the
             * match with the indicated replacement string. If there is no match, returns 
             * the original input string unchanged
            */

            return Regex.Replace(inputText, $@"{pattern}", $@"{replacement}", (RegexOptions)options);
        }

        public string[] Split(string inputText, string pattern, RegexOptionsValues options = 0)
        {
            /*
             * Applies a pattern to the input string, and returns an array of substrings using
             * the pattern to effectively define a delimiter
            */

            return Regex.Split(inputText, $@"{pattern}", (RegexOptions)options);
        }

    }
}
