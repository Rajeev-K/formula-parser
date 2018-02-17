// Copyright (c) 2013-present, Rajeev-K.

using System;
using System.Text;
using System.Globalization;
using Microsoft.VisualBasic;
using System.Reflection;

namespace FormulaParser
{
    public class Util
    {
        public static bool IsValidClsIdentiferFirstChar(char c)
        {
            UnicodeCategory cat = char.GetUnicodeCategory(c);
            switch (cat)
            {
            case UnicodeCategory.LowercaseLetter:
            case UnicodeCategory.UppercaseLetter:
            case UnicodeCategory.TitlecaseLetter:
            case UnicodeCategory.LetterNumber:
            case UnicodeCategory.ModifierLetter:
            case UnicodeCategory.OtherLetter:
                return true;
            default:
                return false;
            }
        }

        public static bool IsValidClsIdentifierSubsequentChar(char c)
        {
            UnicodeCategory cat = char.GetUnicodeCategory(c);
            switch (cat)
            {
            case UnicodeCategory.LowercaseLetter:
            case UnicodeCategory.UppercaseLetter:
            case UnicodeCategory.TitlecaseLetter:
            case UnicodeCategory.LetterNumber:
            case UnicodeCategory.ModifierLetter:
            case UnicodeCategory.OtherLetter:
            case UnicodeCategory.ConnectorPunctuation:
            case UnicodeCategory.DecimalDigitNumber:
            case UnicodeCategory.Format:
            case UnicodeCategory.NonSpacingMark:
            case UnicodeCategory.SpacingCombiningMark:
                return true;
            default:
                return false;
            }
        }

        public static bool IsValidClsIdentifier(string s, out string message)
        {
            if (string.IsNullOrEmpty(s))
            {
                message = "Cannot be empty.";
                return false;
            }
            if (!IsValidClsIdentiferFirstChar(s[0]))
            {
                message = "The first character must be a letter.";
                return false;
            }
            for (int i = 1; i < s.Length; i++)
            {
                if (!IsValidClsIdentifierSubsequentChar(s[i]))
                {
                    message = string.Format("\"{0}\" is not an allowed character.", s[i]);
                    return false;
                }
            }
            message = string.Empty;
            return true;
        }

        public static bool IsPrimitiveType(Type type)
        {
            return (type == typeof(int)) ||
                   (type == typeof(double)) ||
                   (type == typeof(float)) ||
                   (type == typeof(string)) ||
                   (type == typeof(bool)) ||
                   (type == typeof(char));
        }
    }
}
