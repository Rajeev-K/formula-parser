// Copyright (c) 2013-present, Rajeev-K.

using System;
using System.IO;
using System.Globalization;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;

namespace FormulaParser
{
    internal class TokenStreamWithLookahead
    {
        private Scanner m_scanner;
        private List<Token> m_lookAheadBuffer = new List<Token>();

        internal TokenStreamWithLookahead(Scanner scanner)
        {
            m_scanner = scanner;
        }

        internal Token Read()
        {
            if (m_lookAheadBuffer.Count != 0)
            {
                Token t = m_lookAheadBuffer[0];
                m_lookAheadBuffer.RemoveAt(0);
                return t;
            }
            return m_scanner.NextToken();
        }

        internal Token Lookahead(int index)
        {
            while (index >= m_lookAheadBuffer.Count)
            {
                Token t = m_scanner.NextToken();
                m_lookAheadBuffer.Add(t);
            }
            return m_lookAheadBuffer[index];
        }
    }

    public class ExpressionParser
    {
        private TokenStreamWithLookahead ts;

        public ExpressionParser(TextReader input)
        {
            ts = new TokenStreamWithLookahead(new Scanner(input));
        }

        internal Node Parse()
        {
            Node node = ParsePrivate();
            Token t = ts.Lookahead(0);
            if (t.TokenType != TokenType.Eof)
                throw new ParserException(string.Format("Expecting operator but found {0}", t.ToString()), t.SourceIndex);
            return node;
        }

        private Node ParsePrivate()
        {
            return ParseOrExpression();
        }

        private Node ParseOrExpression()
        {
            Node left = ParseAndExpression();
            TokenType tokenType;
            while ((tokenType = ts.Lookahead(0).TokenType) == TokenType.Or || tokenType == TokenType.Xor)
            {
                Token t = ts.Read();
                Node right = ParseAndExpression();
                if (tokenType == TokenType.Or)
                    left = new OrNode(left, right, t.SourceIndex);
                else
                    left = new XorNode(left, right, t.SourceIndex);
            }
            return left;
        }

        private Node ParseAndExpression()
        {
            Node left = ParseNotExpression();
            while (ts.Lookahead(0).TokenType == TokenType.And)
            {
                Token t = ts.Read();
                Node right = ParseNotExpression();
                left = new AndNode(left, right, t.SourceIndex);
            }
            return left;
        }

        private Node ParseNotExpression()
        {
            if (ts.Lookahead(0).TokenType == TokenType.Not)
            {
                Token t = ts.Read();
                Node node = ParseComparisonExpression();
                return new NotNode(node, t.SourceIndex);
            }
            return ParseComparisonExpression();
        }

        private Node ParseComparisonExpression()
        {
            Node left = ParseConcatExpression();
            if (ts.Lookahead(0).TokenType == TokenType.RelationalOperator)
            {
                Token relop = ts.Read();
                Node right = ParseConcatExpression();
                return new ComparisonNode(left, right, (RelationalOperator)relop.Value, relop.SourceIndex);
            }
            else
            {
                return left;
            }
        }

        private Node ParseConcatExpression()
        {
            Node left = ParseSumExpression();
            while (ts.Lookahead(0).TokenType == TokenType.Concatenate)
            {
                Token t = ts.Read();
                Node right = ParseSumExpression();
                left = new ConcatenationNode(left, right, t.SourceIndex);
            }
            return left;
        }

        private Node ParseSumExpression()
        {
            Node left = ParseModExpression();
            TokenType tokenType;
            while ((tokenType = ts.Lookahead(0).TokenType) == TokenType.Plus || tokenType == TokenType.Minus)
            {
                Token t = ts.Read();
                Node right = ParseModExpression();
                if (tokenType == TokenType.Plus)
                    left = new AdditionNode(left, right, t.SourceIndex);
                else
                    left = new SubtractionNode(left, right, t.SourceIndex);
            }
            return left;
        }

        private Node ParseModExpression()
        {
            Node left = ParseIntegerDivisionExpression();
            while (ts.Lookahead(0).TokenType == TokenType.Mod)
            {
                Token t = ts.Read();
                Node right = ParseIntegerDivisionExpression();
                left = new ModNode(left, right, t.SourceIndex);
            }
            return left;
        }

        private Node ParseIntegerDivisionExpression()
        {
            Node left = ParseTermExpression();
            while (ts.Lookahead(0).TokenType == TokenType.IntegerDivide)
            {
                Token t = ts.Read();
                Node right = ParseTermExpression();
                left = new IntegerDivisionNode(left, right, t.SourceIndex);
            }
            return left;
        }

        private Node ParseTermExpression()
        {
            Node left = ParseSignedExpression();
            TokenType tokenType;
            while ((tokenType = ts.Lookahead(0).TokenType) == TokenType.Multiply || tokenType == TokenType.Divide)
            {
                Token t = ts.Read();
                Node right = ParseSignedExpression();
                if (tokenType == TokenType.Multiply)
                    left = new MultiplicationNode(left, right, t.SourceIndex);
                else
                    left = new DivisionNode(left, right, t.SourceIndex);
            }
            return left;
        }

        private Node ParseSignedExpression()
        {
            TokenType tokenType;
            if ((tokenType = ts.Lookahead(0).TokenType) == TokenType.Plus || tokenType == TokenType.Minus)
            {
                Token t = ts.Read();
                Node node = ParseFactorWithSuffix();
                if (tokenType == TokenType.Plus)
                    return node;
                else
                    return new NegationNode(node, t.SourceIndex);
            }
            else
            {
                return ParseFactorWithSuffix();
            }
        }

        private Node ParseFactorWithSuffix()
        {
            bool factorIsNothing = (ts.Lookahead(0).TokenType == TokenType.Nothing);
            Node left = ParseFactor();
            if (factorIsNothing)
                return left;
            while (ts.Lookahead(0).TokenType == TokenType.Dot)
            {
                Token t = ts.Read();
                Node right = ParseIdentifierOrFunction();
                left = new DotNode(left, right, t.SourceIndex);
            }
            return left;
        }

        private Node ParseFactor()
        {
            if (ts.Lookahead(0).TokenType == TokenType.Identifier)
                return ParseIdentifierOrFunction();

            Token token = ts.Read();
            switch (token.TokenType)
            {
            case TokenType.String:
            case TokenType.Integer:
            case TokenType.Double:
            case TokenType.Character:
            case TokenType.Boolean:
            case TokenType.Nothing:
                return new ConstantNode(token, token.SourceIndex);
            case TokenType.OpenParen:
                Node node = ParsePrivate();
                Token t = ts.Read();
                if (t.TokenType != TokenType.CloseParen)
                    throw new ParserException("Unmatched parenthesis", t.SourceIndex);
                return node;
            case TokenType.Eof:
                throw new ParserException("Formula ended unexpectedly", token.SourceIndex);
            default:
                throw new ParserException(string.Format("Unexpected token {0}", token), token.SourceIndex);
            }
        }

        private Node ParseIdentifierOrFunction()
        {
            Token token = ts.Read();
            if (token.TokenType != TokenType.Identifier)
                throw new ParserException(string.Format("Expecting identifier, found {0}", token), token.SourceIndex);

            if (ts.Lookahead(0).TokenType == TokenType.OpenParen)
                return new FunctionCallNode((string)token.Value, ParseArgumentList(), token.SourceIndex);
            else
                return new IdentifierNode((string)token.Value, token.SourceIndex);
        }

        private List<Node> ParseArgumentList()
        {
            Token openParen = ts.Read();
            Debug.Assert(openParen.TokenType == TokenType.OpenParen);

            List<Node> argumentList = new List<Node>();
            TokenType tokenType = ts.Lookahead(0).TokenType;
            if (tokenType == TokenType.CloseParen)
            {
                ts.Read();
                return argumentList;   // argument list is empty
            }
            for (; ; )
            {
                argumentList.Add(ParsePrivate());
                
                Token token = ts.Read();
                if (token.TokenType == TokenType.CloseParen)
                    return argumentList;
                if (token.TokenType != TokenType.Comma)
                    throw new ParserException(string.Format("',' or ')' expected, found {0}", token), token.SourceIndex);
            }
        }
    }
}
