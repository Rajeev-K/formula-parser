using System;
using System.Globalization;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Reflection;

namespace FormulaParser
{
    public interface IResolver
    {
        MethodInfo ResolveFunction(string functionName, Type[] argTypes);
        MethodInfo ResolveMemberFunction(Type declaringType, string functionName, Type[] argTypes);
        MemberInfo ResolveIdentifier(string identifier);
        object GetMemberValue(object obj, string identifier);
        Type GetMemberType(Type type, string identifier);
        Type GetIdentifierType(string identifier);
        object GetIdentifierValue(string identifier);
        string GetCaseSensitiveName(string identifier);
    }

    public abstract class Node
    {
        private int sourceIndex = -1;

        public Node(int sourceIndex)
        {
            this.sourceIndex = sourceIndex;
        }

        public int SourceIndex { get { return sourceIndex; } }

        public abstract void Validate(IResolver resolver);
        public abstract Type GetResultType(IResolver resolver);
        public abstract object Execute(IResolver resolver);
    }

    class OrNode : Node
    {
        private Node left, right;

        public OrNode(Node left, Node right, int sourceIndex) : base(sourceIndex)
        {
            this.left = left;
            this.right = right;
        }

        public override void Validate(IResolver resolver)
        {
            if (left.GetResultType(resolver) != typeof(bool))
                throw new ParserException("type mismatch", SourceIndex);
            if (right.GetResultType(resolver) != typeof(bool))
                throw new ParserException("type mismatch", SourceIndex);
            left.Validate(resolver);
            right.Validate(resolver);
        }

        public override Type GetResultType(IResolver resolver)
        {
            return typeof(bool);
        }

        public override object Execute(IResolver resolver)
        {
            object leftResult = left.Execute(resolver);
            if ((bool)leftResult)
                return true;
            object rightResult = right.Execute(resolver);
            return ((bool)rightResult);
        }
    }

    class XorNode : Node
    {
        private Node left, right;

        public XorNode(Node left, Node right, int sourceIndex) : base(sourceIndex)
        {
            this.left = left;
            this.right = right;
        }

        public override void Validate(IResolver resolver)
        {
            if (left.GetResultType(resolver) != typeof(bool))
                throw new ParserException("type mismatch", SourceIndex);
            if (right.GetResultType(resolver) != typeof(bool))
                throw new ParserException("type mismatch", SourceIndex);
            left.Validate(resolver);
            right.Validate(resolver);
        }

        public override Type GetResultType(IResolver resolver)
        {
            return typeof(bool);
        }

        public override object Execute(IResolver resolver)
        {
            object leftResult = left.Execute(resolver);
            object rightResult = right.Execute(resolver);
            return ((bool)leftResult) ^ ((bool)rightResult);
        }
    }

    class AndNode : Node
    {
        private Node left, right;

        public AndNode(Node left, Node right, int sourceIndex) : base(sourceIndex)
        {
            this.left = left;
            this.right = right;
        }

        public override void Validate(IResolver resolver)
        {
            if (left.GetResultType(resolver) != typeof(bool))
                throw new ParserException("type mismatch", SourceIndex);
            if (right.GetResultType(resolver) != typeof(bool))
                throw new ParserException("type mismatch", SourceIndex);
            left.Validate(resolver);
            right.Validate(resolver);
        }

        public override Type GetResultType(IResolver resolver)
        {
            return typeof(bool);
        }

        public override object Execute(IResolver resolver)
        {
            object leftResult = left.Execute(resolver);
            if (!(bool)leftResult)
                return false;
            object rightResult = right.Execute(resolver);
            return (bool)rightResult;
        }
    }

    class NotNode : Node
    {
        private Node child;

        public NotNode(Node sub, int sourceIndex) : base(sourceIndex)
        {
            this.child = sub;
        }

        public override void Validate(IResolver resolver)
        {
            if (child.GetResultType(resolver) != typeof(bool))
                throw new ParserException("type mismatch", SourceIndex);
            child.Validate(resolver);
        }

        public override Type GetResultType(IResolver resolver)
        {
            return typeof(bool);
        }

        public override object Execute(IResolver resolver)
        {
            object result = child.Execute(resolver);
            return !(bool)result;
        }
    }

    class ComparisonNode : Node
    {
        private Node left, right;
        private RelationalOperator relop;

        public ComparisonNode(Node left, Node right, RelationalOperator relop, int sourceIndex) : base(sourceIndex)
        {
            this.left = left;
            this.right = right;
            this.relop = relop;
        }

        public override void Validate(IResolver resolver)
        {
            if (relop != RelationalOperator.Is)
            {
                if (!Util.IsPrimitiveType(left.GetResultType(resolver)))
                    throw new ParserException("type mismatch", SourceIndex);
                if (!Util.IsPrimitiveType(right.GetResultType(resolver)))
                    throw new ParserException("type mismatch", SourceIndex);
            }
            left.Validate(resolver);
            right.Validate(resolver);
        }

        public override Type GetResultType(IResolver resolver)
        {
            return typeof(bool);
        }

        public override object Execute(IResolver resolver)
        {
            object leftResult = left.Execute(resolver);
            object rightResult = right.Execute(resolver);

            if (relop == RelationalOperator.Is)
                return leftResult == rightResult;

            int k;
            if (leftResult is string || rightResult is string)
                k = Convert.ToString(leftResult).CompareTo(Convert.ToString(rightResult));
            else if (leftResult is double || rightResult is double)
                k = Convert.ToDouble(leftResult).CompareTo(Convert.ToDouble(rightResult));
            else if (leftResult is int || rightResult is int)
                k = Convert.ToInt32(leftResult).CompareTo(Convert.ToInt32(rightResult));
            else
                k = ((IComparable)leftResult).CompareTo(rightResult);
            switch (relop)
            {
            case RelationalOperator.Equals: return (k == 0);
            case RelationalOperator.GreaterThan: return (k > 0);
            case RelationalOperator.GreaterThanEquals: return (k >= 0);
            case RelationalOperator.LessThan: return (k < 0);
            case RelationalOperator.LessThanEquals: return (k <= 0);
            case RelationalOperator.NotEquals: return (k != 0);
            }
            Debug.Assert(false);
            throw new ParserException("Internal error: Invalid relational operator", SourceIndex);
        }
    }

    class ConcatenationNode : Node
    {
        private Node left, right;

        public ConcatenationNode(Node left, Node right, int sourceIndex) : base(sourceIndex)
        {
            this.left = left;
            this.right = right;
        }

        public override void Validate(IResolver resolver)
        {
            left.Validate(resolver);
            right.Validate(resolver);
        }

        public override Type GetResultType(IResolver resolver)
        {
            return typeof(string);
        }

        public override object Execute(IResolver resolver)
        {
            object leftResult = left.Execute(resolver);
            object rightResult = right.Execute(resolver);
            string l = (leftResult == null) ? string.Empty : leftResult.ToString();
            string r = (rightResult == null) ? string.Empty : rightResult.ToString();
            return l + r;
        }
    }

    class AdditionNode : Node
    {
        private Node left, right;

        public AdditionNode(Node left, Node right, int sourceIndex) : base(sourceIndex)
        {
            this.left = left;
            this.right = right;
        }

        public override void Validate(IResolver resolver)
        {
            Type leftType = left.GetResultType(resolver);
            if (leftType != typeof(Int32) && leftType != typeof(double) && leftType != typeof(string))
                throw new ParserException("type mismatch", SourceIndex);
            Type rightType = right.GetResultType(resolver);
            if (rightType != typeof(Int32) && rightType != typeof(double) && rightType != typeof(string))
                throw new ParserException("type mismatch", SourceIndex);
            left.Validate(resolver);
            right.Validate(resolver);
        }

        public override Type GetResultType(IResolver resolver)
        {
            Type leftType = left.GetResultType(resolver);
            if (leftType == typeof(string))
                return typeof(string);

            Type rightType = right.GetResultType(resolver);
            if (rightType == typeof(string))
                return typeof(string);

            if (leftType == typeof(double) || leftType == typeof(float) ||
                   rightType == typeof(double) || rightType == typeof(float))
                return typeof(double);
            else
                return typeof(int);
        }

        public override object Execute(IResolver resolver)
        {
            object leftResult = left.Execute(resolver);
            object rightResult = right.Execute(resolver);
            if (leftResult is string || rightResult is string)
                return leftResult.ToString() + rightResult.ToString();
            if (leftResult is double || leftResult is float || rightResult is double || rightResult is float)
                return Convert.ToDouble(leftResult) + Convert.ToDouble(rightResult);
            return Convert.ToInt32(leftResult) + Convert.ToInt32(rightResult);
        }
    }

    class SubtractionNode : Node
    {
        private Node left, right;

        public SubtractionNode(Node left, Node right, int sourceIndex) : base(sourceIndex)
        {
            this.left = left;
            this.right = right;
        }

        public override void Validate(IResolver resolver)
        {
            Type leftType = left.GetResultType(resolver);
            if (leftType != typeof(Int32) && leftType != typeof(double))
                throw new ParserException("type mismatch", SourceIndex);
            Type rightType = right.GetResultType(resolver);
            if (rightType != typeof(Int32) && rightType != typeof(double))
                throw new ParserException("type mismatch", SourceIndex);
            left.Validate(resolver);
            right.Validate(resolver);
        }

        public override Type GetResultType(IResolver resolver)
        {
            Type leftType = left.GetResultType(resolver);
            Type rightType = right.GetResultType(resolver);
            if (leftType == typeof(double) || leftType == typeof(float) ||
                 rightType == typeof(double) || rightType == typeof(float))
                return typeof(double);
            return typeof(int);
        }

        public override object Execute(IResolver resolver)
        {
            object leftResult = left.Execute(resolver);
            object rightResult = right.Execute(resolver);
            if (leftResult is double || leftResult is float || 
                 rightResult is double || rightResult is float)
                return Convert.ToDouble(leftResult) - Convert.ToDouble(rightResult);
            return Convert.ToInt32(leftResult) - Convert.ToInt32(rightResult);
        }
    }

    class ModNode : Node
    {
        private Node left, right;

        public ModNode(Node left, Node right, int sourceIndex) : base(sourceIndex)
        {
            this.left = left;
            this.right = right;
        }

        public override void Validate(IResolver resolver)
        {
            Type leftType = left.GetResultType(resolver);
            if (leftType != typeof(Int32))
                throw new ParserException("type mismatch", SourceIndex);
            Type rightType = right.GetResultType(resolver);
            if (rightType != typeof(Int32))
                throw new ParserException("type mismatch", SourceIndex);
            left.Validate(resolver);
            right.Validate(resolver);
        }

        public override Type GetResultType(IResolver resolver)
        {
            return typeof(int);
        }

        public override object Execute(IResolver resolver)
        {
            object leftResult = left.Execute(resolver);
            object rightResult = right.Execute(resolver);
            return (int)leftResult % (int)rightResult;
        }
    }

    class IntegerDivisionNode : Node
    {
        private Node left, right;

        public IntegerDivisionNode(Node left, Node right, int sourceIndex) : base(sourceIndex)
        {
            this.left = left;
            this.right = right;
        }

        public override void Validate(IResolver resolver)
        {
            Type leftType = left.GetResultType(resolver);
            if (leftType != typeof(Int32))
                throw new ParserException("type mismatch", SourceIndex);
            Type rightType = right.GetResultType(resolver);
            if (rightType != typeof(Int32))
                throw new ParserException("type mismatch", SourceIndex);
            left.Validate(resolver);
            right.Validate(resolver);
        }

        public override Type GetResultType(IResolver resolver)
        {
            return typeof(int);
        }

        public override object Execute(IResolver resolver)
        {
            object leftResult = left.Execute(resolver);
            object rightResult = right.Execute(resolver);
            return (int)leftResult / (int)rightResult;
        }
    }

    class MultiplicationNode : Node
    {
        private Node left, right;

        public MultiplicationNode(Node left, Node right, int sourceIndex) : base(sourceIndex)
        {
            this.left = left;
            this.right = right;
        }

        public override void Validate(IResolver resolver)
        {
            Type leftType = left.GetResultType(resolver);
            if (leftType != typeof(Int32) && leftType != typeof(double))
                throw new ParserException("type mismatch", SourceIndex);
            Type rightType = right.GetResultType(resolver);
            if (rightType != typeof(Int32) && rightType != typeof(double))
                throw new ParserException("type mismatch", SourceIndex);
            left.Validate(resolver);
            right.Validate(resolver);
        }

        public override Type GetResultType(IResolver resolver)
        {
            if (left.GetResultType(resolver) == typeof(int) && right.GetResultType(resolver) == typeof(int))
                return typeof(int);
            return typeof(double);
        }

        public override object Execute(IResolver resolver)
        {
            object leftResult = left.Execute(resolver);
            object rightResult = right.Execute(resolver);
            if (leftResult is int && rightResult is int)
                return (int)leftResult * (int)rightResult;
            return Convert.ToDouble(leftResult) * Convert.ToDouble(rightResult);
        }
    }

    class DivisionNode : Node
    {
        private Node left, right;

        public DivisionNode(Node left, Node right, int sourceIndex) : base(sourceIndex)
        {
            this.left = left;
            this.right = right;
        }

        public override void Validate(IResolver resolver)
        {
            Type leftType = left.GetResultType(resolver);
            if (leftType != typeof(Int32) && leftType != typeof(double))
                throw new ParserException("type mismatch", SourceIndex);
            Type rightType = right.GetResultType(resolver);
            if (rightType != typeof(Int32) && rightType != typeof(double))
                throw new ParserException("type mismatch", SourceIndex);
            left.Validate(resolver);
            right.Validate(resolver);
        }

        public override Type GetResultType(IResolver resolver)
        {
            return typeof(double);
        }

        public override object Execute(IResolver resolver)
        {
            object leftResult = left.Execute(resolver);
            object rightResult = right.Execute(resolver);
            return Convert.ToDouble(leftResult) / Convert.ToDouble(rightResult);
        }
    }

    class NegationNode : Node
    {
        private Node child;

        public NegationNode(Node child, int sourceIndex) : base(sourceIndex)
        {
            this.child = child;
        }

        public override void Validate(IResolver resolver)
        {
            Type childType = child.GetResultType(resolver);
            if (childType != typeof(Int32) && childType != typeof(double))
                throw new ParserException("type mismatch", SourceIndex);
            child.Validate(resolver);
        }

        public override Type GetResultType(IResolver resolver)
        {
            if (child.GetResultType(resolver) == typeof(int))
                return typeof(int);
            return typeof(double);
        }

        public override object Execute(IResolver resolver)
        {
            object result = child.Execute(resolver);
            if (result is int)
                return -1 * (int)result;
            else
                return -1 * (double)result;
        }
    }

    class ConstantNode : Node
    {
        private Token value;

        public ConstantNode(Token value, int sourceIndex) : base(sourceIndex)
        {
            this.value = value;
        }

        public override object Execute(IResolver resolver)
        {
            return value.Value;
        }

        public override void Validate(IResolver resolver)
        {
        }

        public override Type GetResultType(IResolver resolver)
        {
            switch (value.TokenType)
            {
            case TokenType.String:
                return typeof(string);
            case TokenType.Character:
                return typeof(char);
            case TokenType.Boolean:
                return typeof(bool);
            case TokenType.Integer:
                return typeof(int);
            case TokenType.Double:
                return typeof(double);
            case TokenType.Nothing:
                return typeof(object);
            default:
                Debug.Assert(false, "invalid constant type");
                throw new ParserException("Internal error: Invalid constant", SourceIndex);
            }
        }
    }

    class DotNode : Node
    {
        private Node left, right;

        public DotNode(Node left, Node right, int sourceIndex) : base(sourceIndex)
        {
            this.left = left;
            this.right = right;
        }

        public override void Validate(IResolver resolver)
        {
            left.Validate(resolver);
            Type leftType = left.GetResultType(resolver);

            IdentifierNode identifierNode = right as IdentifierNode;
            if (identifierNode != null)
            {
                resolver.GetMemberType(leftType, identifierNode.Identifier);
                return;
            }

            FunctionCallNode functionNode = right as FunctionCallNode;
            if (functionNode != null)
            {
                functionNode.DeclaringType = left.GetResultType(resolver);
                return;
            }
        }

        public override Type GetResultType(IResolver resolver)
        {
            Type leftType = left.GetResultType(resolver);
            IdentifierNode identifierNode = right as IdentifierNode;
            if (identifierNode != null)
            {
                return resolver.GetMemberType(leftType, identifierNode.Identifier);
            }
            else
            {
                FunctionCallNode functionNode = (FunctionCallNode)right;
                functionNode.DeclaringType = leftType;
                return functionNode.GetResultType(resolver);
            }
        }

        public override object Execute(IResolver resolver)
        {
            object leftResult = left.Execute(resolver);
            if (leftResult == null)
            {
                throw new Exception(String.Format("Formula error: Cannot get value of {0}.{1} because {0} is Null.",
                         left.ToString(), right.ToString()));
            }
            if (right is IdentifierNode)
            {
                return resolver.GetMemberValue(leftResult, ((IdentifierNode)right).Identifier);
            }
            else
            {
                FunctionCallNode functionNode = (FunctionCallNode)right;
                functionNode.DeclaringType = leftResult.GetType();
                return functionNode.Execute(leftResult, resolver);
            }
        }
    }

    /// <summary>
    /// Represents a variable, property, class or enum type.
    /// </summary>
    class IdentifierNode : Node
    {
        private string identifier;
        private bool resolved;
        PropertyInfo propertyInfo;
        Type classOrEnumType;

        public IdentifierNode(string identifier, int sourceIndex) : base(sourceIndex)
        {
            this.identifier = identifier;
        }

        public string Identifier
        {
            get { return identifier; }
        }

        private void Resolve(IResolver resolver)
        {
            MemberInfo memberInfo = resolver.ResolveIdentifier(identifier);
            if (memberInfo != null)
            {
                if ((propertyInfo = memberInfo as PropertyInfo) == null)
                    classOrEnumType = memberInfo as Type;
            }
            resolved = true;
        }

        public override void Validate(IResolver resolver)
        {
            GetResultType(resolver);
        }

        public override Type GetResultType(IResolver resolver)
        {
            if (!resolved)
                Resolve(resolver);
            if (propertyInfo != null)
                return propertyInfo.PropertyType;
            else if (classOrEnumType != null)
                return classOrEnumType;
            else
                return resolver.GetIdentifierType(identifier);
        }

        public override object Execute(IResolver resolver)
        {
            if (!resolved)
                Resolve(resolver);
            if (propertyInfo != null)
                return propertyInfo.GetValue(null, null);
            else if (classOrEnumType != null)
                return classOrEnumType;
            else
                return resolver.GetIdentifierValue(identifier);
        }

        public override string ToString()
        {
            return identifier;
        }
    }

    class FunctionCallNode : Node
    {
        private string functionName;
        private List<Node> arguments;
        private MethodInfo methodInfo;
        private Type declaringType;

        public FunctionCallNode(string functionName, List<Node> arguments, int sourceIndex) : base(sourceIndex)
        {
            this.functionName = functionName;
            this.arguments = arguments;
        }

        public Type DeclaringType
        {
            set { declaringType = value; }
        }

        private void Resolve(IResolver resolver)
        {
            Type[] argTypes = new Type[arguments.Count];
            for (int i = 0; i < arguments.Count; i++)
                argTypes[i] = arguments[i].GetResultType(resolver);
            if (declaringType == null)
                methodInfo = resolver.ResolveFunction(functionName, argTypes);
            else
                methodInfo = resolver.ResolveMemberFunction(declaringType, functionName, argTypes);
        }

        public override void Validate(IResolver resolver)
        {
            for (int i = 0; i < arguments.Count; i++)
                arguments[i].Validate(resolver);
            if (methodInfo == null)
                Resolve(resolver);
        }

        public override Type GetResultType(IResolver resolver)
        {
            if (methodInfo == null)
                Resolve(resolver);
            return methodInfo.ReturnType;
        }

        public override object Execute(IResolver resolver)
        {
            return Execute(null, resolver);
        }

        public object Execute(object instance, IResolver resolver)
        {
            if (methodInfo == null)
                Resolve(resolver);
            object[] values = new object[arguments.Count];
            for (int i = 0; i < arguments.Count; i++)
                values[i] = arguments[i].Execute(resolver);
            return methodInfo.Invoke(instance, values);
        }
    }
}
