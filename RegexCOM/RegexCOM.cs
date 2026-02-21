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
            return Regex.IsMatch(inputText, $@"{pattern}", (RegexOptions)options);
        }

        public string Match(string inputText, string pattern, RegexOptionsValues options = 0)
        {
            return Regex.Match(inputText, $@"{pattern}", (RegexOptions)options).ToString();
        }

        public string[] Matches(string inputText, string pattern, RegexOptionsValues options = 0)
        {
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
            return Regex.Replace(inputText, $@"{pattern}", $@"{replacement}", (RegexOptions)options);
        }

        public string[] Split(string inputText, string pattern, RegexOptionsValues options = 0)
        {
            return Regex.Split(inputText, $@"{pattern}", (RegexOptions)options);
        }

    }
}
