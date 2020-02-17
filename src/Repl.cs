using System;
using System.IO;
using System.Reflection;

namespace FormulaParser
{
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
