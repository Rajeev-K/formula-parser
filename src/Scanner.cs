// Copyright (c) 2013-present, Rajeev-K.

using System;
using System.Globalization;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Diagnostics;

namespace FormulaParser
{
    enum TokenType
    {
        Integer, Double, String, Boolean, Character, Nothing,   // primitives
        Plus, Minus, Multiply, Divide, IntegerDivide, Mod,      // aritmetic + - * / \ Mod
        Comma, Dot,                                             // comma, dot
        OpenParen, CloseParen,                                  // ( )
        Identifier,                                             // identifier
        RelationalOperator,                                     // = > < >= <= <> is
        Not, And, Or, Xor,                                      // boolean operators
        Concatenate,                                            // &
        Eof,
    };

    enum RelationalOperator { Equals, GreaterThan, LessThan, GreaterThanEquals, LessThanEquals, NotEquals, Is };

    class Token
    {
        private TokenType m_tokenType;
        private object m_value;
        private int m_sourceIndex;
        private int m_sourceLength;

        public Token(TokenType tokenType, object value, int sourceIndex, int sourceLength)
            : this(tokenType, sourceIndex, sourceLength)
        {
            m_value = value;
        }

        public Token(TokenType tokenType, int sourceIndex, int sourceLength)
        {
            m_tokenType = tokenType;
            m_sourceIndex = sourceIndex;
            m_sourceLength = sourceLength;
        }

        public int SourceIndex
        {
            get { return m_sourceIndex; }
        }

        public int SourceLength
        {
            get { return m_sourceLength; }
        }

        internal TokenType TokenType
        {
            get { return m_tokenType; }
        }

        internal object Value
        {
            get { return m_value; }
        }

        public override string ToString()
        {
            switch (m_tokenType)
            {
            case TokenType.And: return "AND";
            case TokenType.Boolean: return m_value.ToString();
            case TokenType.Character: return "\"" + m_value.ToString() + "\"c";
            case TokenType.CloseParen: return ")";
            case TokenType.Comma: return ",";
            case TokenType.Dot: return ".";
            case TokenType.Concatenate: return "&";
            case TokenType.Divide: return "/";
            case TokenType.Double: return m_value.ToString();
            case TokenType.Eof: return "end-of-line";
            case TokenType.Identifier: return m_value.ToString();
            case TokenType.Integer: return m_value.ToString();
            case TokenType.IntegerDivide: return "\\";
            case TokenType.Minus: return "-";
            case TokenType.Mod: return "MOD";
            case TokenType.Multiply: return "*";
            case TokenType.Not: return "NOT";
            case TokenType.OpenParen: return "(";
            case TokenType.Or: return "OR";
            case TokenType.Xor: return "XOR";
            case TokenType.Plus: return "+";
            case TokenType.String: return "\"" + m_value.ToString() + "\"";
            case TokenType.RelationalOperator:
                switch ((RelationalOperator)m_value)
                {
                case RelationalOperator.Equals: return "=";
                case RelationalOperator.GreaterThan: return ">";
                case RelationalOperator.GreaterThanEquals: return ">=";
                case RelationalOperator.LessThan: return "<";
                case RelationalOperator.LessThanEquals: return "<=";
                case RelationalOperator.NotEquals: return "<>";
                case RelationalOperator.Is: return "is";
                }
                break;
            }
            Debug.Assert(false, string.Format("unknown token {0}", m_tokenType));
            return m_tokenType.ToString();
        }
    }

    class Scanner
    {
        private TextReaderWithLookahead m_input;
        private const int EOF = -1;

        public Scanner(TextReader input)
        {
            m_input = new TextReaderWithLookahead(input);
        }

        private void SkipWhite()
        {
            for (; ; )
            {
                int c = m_input.Lookahead(0);
                if (c == EOF)
                    return;
                if (!char.IsWhiteSpace((char)c))
                    break;
                m_input.Read();
            }
        }

        public Token NextToken()
        {
            SkipWhite();
            int sourceIndex = m_input.Offset;
            int n = m_input.Lookahead(0);
            if (n == EOF)
                return new Token(TokenType.Eof, sourceIndex, 0);
            char c = (char)n;

            switch (c)
            {
            case '+':
                m_input.Read();
                return new Token(TokenType.Plus, sourceIndex, 1);
            case '-':
                m_input.Read();
                return new Token(TokenType.Minus, sourceIndex, 1);
            case '*':
                m_input.Read();
                return new Token(TokenType.Multiply, sourceIndex, 1);
            case '/':
                m_input.Read();
                return new Token(TokenType.Divide, sourceIndex, 1);
            case '\\':
                m_input.Read();
                return new Token(TokenType.IntegerDivide, sourceIndex, 1);
            case ',':
                m_input.Read();
                return new Token(TokenType.Comma, sourceIndex, 1);
            case '(':
                m_input.Read();
                return new Token(TokenType.OpenParen, sourceIndex, 1);
            case ')':
                m_input.Read();
                return new Token(TokenType.CloseParen, sourceIndex, 1);
            case '=':
                m_input.Read();
                return new Token(TokenType.RelationalOperator, RelationalOperator.Equals, sourceIndex, 1);
            case '&':
                m_input.Read();
                return new Token(TokenType.Concatenate, sourceIndex, 1);
            case '<':
                {
                    m_input.Read();
                    char c2 = (char)m_input.Lookahead(0);
                    if (c2 == '=')
                    {
                        m_input.Read();
                        return new Token(TokenType.RelationalOperator, RelationalOperator.LessThanEquals, sourceIndex, 2);
                    }
                    else if (c2 == '>')
                    {
                        m_input.Read();
                        return new Token(TokenType.RelationalOperator, RelationalOperator.NotEquals, sourceIndex, 2);
                    }
                }
                return new Token(TokenType.RelationalOperator, RelationalOperator.LessThan, sourceIndex, 1);
            case '>':
                {
                    m_input.Read();
                    if ((char)m_input.Lookahead(0) == '=')
                    {
                        m_input.Read();
                        return new Token(TokenType.RelationalOperator, RelationalOperator.GreaterThanEquals, sourceIndex, 2);
                    }
                }
                return new Token(TokenType.RelationalOperator, RelationalOperator.GreaterThan, sourceIndex, 1);
            case '.':
                if (!char.IsDigit((char)m_input.Lookahead(1)))
                {
                    m_input.Read();
                    return new Token(TokenType.Dot, sourceIndex, 1);
                }
                break;
            }

            // String or character
            // todo: handle escape: treat "" as "
            if (c == '"')
            {
                m_input.Read();   // '"'
                StringBuilder sb = new StringBuilder();
                for (; ; )
                {
                    int m = m_input.Lookahead(0);
                    if (m == EOF)
                        throw new ParserException("String constants must end with a double quote", m_input.Offset);
                    char c2 = (char)m;
                    if (c2 == '"')
                        break;
                    sb.Append(c2);
                    m_input.Read();
                }
                m_input.Read();   // Read '"'
                if ((char)m_input.Lookahead(0) == 'c')
                {
                    // A character can be written as "x"c
                    if (sb.Length != 1)
                        throw new ParserException("Character constant must contain exactly one character", m_input.Offset);
                    m_input.Read();
                    return new Token(TokenType.Character, sb[0], sourceIndex, m_input.Offset - sourceIndex);
                }
                return new Token(TokenType.String, sb.ToString(), sourceIndex, m_input.Offset - sourceIndex);
            }

            // Number
            else if (char.IsDigit(c) || c == '.')
            {
                StringBuilder sb = new StringBuilder();
                while (char.IsDigit(c))
                {
                    m_input.Read();
                    sb.Append(c);
                    int m = m_input.Lookahead(0);
                    if (m == EOF)
                        break;
                    c = (char)m;
                }
                // At this point we are at EOF or a non-digit.
                if (c != '.' && c != 'e')
                {
                    // We have at least one digit at this point.
                    int val;
                    if (Int32.TryParse(sb.ToString(), out val))
                        return new Token(TokenType.Integer, val, sourceIndex, m_input.Offset - sourceIndex);
                    else
                        throw new ParserException(string.Format("Invalid number {0}", sb.ToString()), m_input.Offset);
                }
                // Get fraction part if present
                if (c == '.' && char.IsDigit((char)m_input.Lookahead(1)))
                {
                    m_input.Read();   // '.'
                    sb.Append('.');
                    c = (char)m_input.Lookahead(0);
                    Debug.Assert(char.IsDigit(c));
                    while (char.IsDigit(c))
                    {
                        m_input.Read();
                        sb.Append(c);
                        int m = m_input.Lookahead(0);
                        if (m == EOF)
                            break;
                        c = (char)m;
                    }
                }
                // Get exponent if present
                if (c == 'e')
                {
                    m_input.Read();   // 'e'
                    sb.Append('e');
                    int t = m_input.Lookahead(0);
                    if (t == EOF)
                        throw new ParserException("Invalid exponent", m_input.Offset);
                    c = (char)t;
                    if (c == '+' || c == '-')
                    {
                        m_input.Read();
                        sb.Append(c);
                        t = m_input.Lookahead(0);
                        if (t == EOF)
                            throw new ParserException("Invalid exponent", m_input.Offset);
                        c = (char)t;
                    }
                    if (!char.IsDigit(c))
                        throw new ParserException("Invalid exponent", m_input.Offset);
                    while (char.IsDigit(c))
                    {
                        m_input.Read();
                        sb.Append(c);
                        int m = m_input.Lookahead(0);
                        if (m == EOF)
                            break;
                        c = (char)m;
                    }
                }
                double d;
                if (double.TryParse(sb.ToString(), out d))
                    return new Token(TokenType.Double, d, sourceIndex, m_input.Offset - sourceIndex);
                else
                    throw new ParserException(string.Format("Invalid number {0}", sb.ToString()), m_input.Offset);
            }

            // Keyword or identifier
            else if (FormulaParser.Util.IsValidClsIdentiferFirstChar(c))
            {
                StringBuilder sb = new StringBuilder();
                while (FormulaParser.Util.IsValidClsIdentifierSubsequentChar(c))
                {
                    m_input.Read();
                    sb.Append(c);
                    int u = m_input.Lookahead(0);
                    if (u == EOF)
                        break;
                    c = (char)u;
                }
                string w = sb.ToString();
                switch (w.ToLower())
                {
                case "true": return new Token(TokenType.Boolean, true, sourceIndex, m_input.Offset - sourceIndex);
                case "false": return new Token(TokenType.Boolean, false, sourceIndex, m_input.Offset - sourceIndex);
                case "mod": return new Token(TokenType.Mod, sourceIndex, m_input.Offset - sourceIndex);
                case "not": return new Token(TokenType.Not, sourceIndex, m_input.Offset - sourceIndex);
                case "and": return new Token(TokenType.And, sourceIndex, m_input.Offset - sourceIndex);
                case "or": return new Token(TokenType.Or, sourceIndex, m_input.Offset - sourceIndex);
                case "xor": return new Token(TokenType.Xor, sourceIndex, m_input.Offset - sourceIndex);
                case "is": return new Token(TokenType.RelationalOperator, RelationalOperator.Is, sourceIndex, m_input.Offset - sourceIndex);
                case "nothing": return new Token(TokenType.Nothing, sourceIndex, m_input.Offset - sourceIndex);
                }
                return new Token(TokenType.Identifier, w, sourceIndex, m_input.Offset - sourceIndex);
            }

            throw new ParserException(string.Format("Invalid character {0} in formula", c), m_input.Offset);
        }
    }

    internal class TextReaderWithLookahead
    {
        private TextReader m_streamReader;
        private List<int> m_lookAheadBuffer = new List<int>();
        private int m_offset = 0;

        internal TextReaderWithLookahead(TextReader streamReader)
        {
            m_streamReader = streamReader;
        }

        internal int Read()
        {
            int b;
            if (m_lookAheadBuffer.Count == 0)
            {
                b = m_streamReader.Read();
            }
            else
            {
                b = m_lookAheadBuffer[0];
                m_lookAheadBuffer.RemoveAt(0);
            }
            if (b != -1)
                m_offset++;
            return b;
        }

        internal int Lookahead(int index)
        {
            while (index >= m_lookAheadBuffer.Count)
            {
                int b = m_streamReader.Read();
                m_lookAheadBuffer.Add(b);
            }
            return m_lookAheadBuffer[index];
        }

        internal int Offset
        {
            get { return m_offset; }
        }
    }

    class ParserException : ApplicationException
    {
        int m_sourceIndex = -1;

        public ParserException(string s)
            : base(s)
        {
        }

        public ParserException(string s, int sourceIndex)
            : base(s)
        {
            m_sourceIndex = sourceIndex;
        }

        public int SourceIndex
        {
            get { return m_sourceIndex; }
        }
    }
}
