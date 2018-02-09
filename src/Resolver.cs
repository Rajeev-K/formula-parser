using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;

namespace FormulaParser
{
    class Resolver : IResolver
    {
        private Type[] classesToSearchIn = 
            {
                typeof(FormulaParser.Functions),
            };

        private Type FindType(string typeName)
        {
            Type type;

            Assembly formulaParserAssembly = Assembly.GetExecutingAssembly();
            if ((type = formulaParserAssembly.GetType("FormulaParser." + typeName, false, true)) != null)
                return type;

            return null;
        }

        private PropertyInfo FindProperty(string propertyName)
        {
            PropertyInfo pi = null;
            foreach (Type type in classesToSearchIn)
            {
                pi = type.GetProperty(propertyName, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Static);
                if (pi != null)
                    return pi;
            }
            return null;
        }

        private MethodInfo FindMethod(string functionName, Type[] argTypes)
        {
            MethodInfo mi = null;
            foreach (Type type in classesToSearchIn)
            {
                mi = type.GetMethod(functionName,
                                    BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Static,
                                    (Binder)null, argTypes, null);
                if (mi != null)
                    break;
            }
            return mi;
        }

        /// <summary>
        /// Checks whether a function by supplied name is known, ignoring parameters. Use for diagnostic messages.
        /// </summary>
        private bool DoesFunctionExist(string functionName)
        {
            foreach (Type type in classesToSearchIn)
            {
                MethodInfo[] methods = type.GetMethods(BindingFlags.Public | BindingFlags.Static);
                foreach (MethodInfo method in methods)
                {
                    if (method.Name.Equals(functionName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            return false;
        }

        public virtual string GetCaseSensitiveName(string identifier)
        {
            throw new ParserException(string.Format("Identifier {0} is unknown", identifier));
        }

        public virtual Type GetIdentifierType(string identifier)
        {
            throw new ParserException(string.Format("Identifier {0} is unknown", identifier));
        }

        public virtual object GetIdentifierValue(string identifier)
        {
            throw new ParserException(string.Format("Identifier {0} is unknown", identifier));
        }

        public MemberInfo ResolveIdentifier(string identifier)
        {
            PropertyInfo pi = FindProperty(identifier);
            if (pi != null)
                return pi;
            Type type = FindType(identifier);
            if (type != null)
                return type;
            return null;
        }

        /// <summary>
        /// Checks whether a member function by supplied name is known, ignoring parameters. Use for diagnostic messages.
        /// </summary>
        private bool DoesMemberFunctionExist(Type declaringType, string functionName)
        {
            MethodInfo[] methods = declaringType.GetMethods(BindingFlags.Public | BindingFlags.Instance);
            foreach (MethodInfo method in methods)
            {
                if (method.Name.Equals(functionName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        private bool AllowAccessOfMemberFunction(Type type, string memberName)
        {
            if (type != typeof(DateTime))
                return false;
            return true;
        }

        public MethodInfo ResolveMemberFunction(Type declaringType, string functionName, Type[] argTypes)
        {
            if (!AllowAccessOfMemberFunction(declaringType, functionName))
                throw new ParserException(string.Format("Cannot access specified member of {0}", declaringType.Name));

            MethodInfo mi = declaringType.GetMethod(functionName,
                                BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance,
                                (Binder)null, argTypes, null);
            if (mi != null)
                return mi;
            if (DoesMemberFunctionExist(declaringType, functionName))
            {
                throw new ParserException(string.Format("{0}.{1}: argument count or types do not match",
                                 declaringType.Name, functionName));
            }
            else
            {
                throw new ParserException(string.Format("{0} is not a member of {1}",
                                 functionName, declaringType.Name));
            }
        }

        public MethodInfo ResolveFunction(string functionName, Type[] argTypes)
        {
            MethodInfo mi = FindMethod(functionName, argTypes);
            if (mi != null)
                return mi;
            if (DoesFunctionExist(functionName))
                throw new ParserException(string.Format("{0}: argument count or types do not match", functionName));
            else
                throw new ParserException(string.Format("{0} is not a known function", functionName));
        }

        private bool AllowAccessOfMemberProperty(Type type, string memberName)
        {
            if (type != typeof(DateTime) && 
                type != typeof(FormulaParser.FirstDayOfWeek) &&
                type != typeof(FormulaParser.DateInterval))
                return false;
            return true;
        }

        public Type GetMemberType(Type type, string memberName)
        {
            if (!AllowAccessOfMemberProperty(type, memberName))
                throw new ParserException(string.Format("Cannot access specified member of {0}", type.Name));
            MemberInfo[] members = type.GetMember(memberName,
                       BindingFlags.Public | BindingFlags.Static | BindingFlags.Instance | BindingFlags.IgnoreCase);
            if (members.Length > 0)
            {
                if (members[0] is FieldInfo)
                    return ((FieldInfo)members[0]).FieldType;
                else if (members[0] is PropertyInfo)
                    return ((PropertyInfo)members[0]).PropertyType;
            }
            throw new ParserException(string.Format("No member by the name {0} found in {1}", memberName, type.Name));
        }

        public object GetMemberValue(object obj, string memberName)
        {
            Type type = obj as Type;
            if (type == null)
                type = obj.GetType();
            if (!AllowAccessOfMemberProperty(type, memberName))
                throw new ParserException(string.Format("Cannot access specified member of {0}", type.Name));
            MemberInfo[] members = type.GetMember(memberName, 
                       BindingFlags.Public | BindingFlags.Static | BindingFlags.Instance | BindingFlags.IgnoreCase);
            if (members.Length > 0)
            {
                if (members[0] is FieldInfo)
                    return ((FieldInfo)members[0]).GetValue(obj);
                else if (members[0] is PropertyInfo)
                    return ((PropertyInfo)members[0]).GetValue(obj, null);
            }
            throw new ParserException(string.Format("No member by the name {0} found in {1}", memberName, type.Name));
        }
    }

#if FORMULA_MAIN
    class Program
    {
        static void Main(string[] args)
        {
            IResolver resolver = new Resolver();
            for (; ; )
            {
                string s = Console.ReadLine();
                ExpressionParser parser = new ExpressionParser(new StringReader(s));
                try
                {
                    Node root = parser.Parse();
                    root.Validate(resolver);
                    object result = root.Execute(resolver);
                    Console.WriteLine("= " + result.ToString());
                }
                catch (InvalidCastException)
                {
                    Console.WriteLine("type mismatch");
                }
                catch (FormatException)   // thrown if Convert.ToDouble("abcd")
                {
                    Console.WriteLine("type mismatch");
                }
                catch (ParserException ex)
                {
                    if (ex.SourceIndex >= 0)
                    {
                        for (int i = 0; i < ex.SourceIndex; i++)
                            Console.Write(' ');
                        Console.WriteLine('^');
                    }
                    Console.WriteLine(ex.Message);
                }
                catch (TargetInvocationException ex)
                {
                    Console.WriteLine(ex.InnerException.Message);
                }
            }
        }
    }
#endif
}
