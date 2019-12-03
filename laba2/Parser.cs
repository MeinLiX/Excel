using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace laba2
{
    public class Parser
    {
        public Parser()
        {

        }
        public string str_error = "";
        private List<string> operators = new List<string>(new string[] { "(", ")", "+", "-", "*", "/", "^",
            "max", "min", "inc", "dec","mmax", "mmin" });
        private List<string> standart_operators =
            new List<string>(new string[] { "(", ")", "+", "-", "*", "/", "^" });

        private IEnumerable<string> Separate(string input)
        {
            bool flag_min = false;
            bool flag_max = false;
            int pos = 0;
            while (pos < input.Length)
            {
                while (input[pos] == ' ') pos++;
                string s = string.Empty + input[pos];
                if (!standart_operators.Contains(input[pos].ToString()))
                {
                    if (Char.IsDigit(input[pos]))
                    {
                        for (int i = pos + 1; i < input.Length && (Char.IsDigit(input[i])); i++)
                            s += input[i];
                    }
                    else if (Char.IsLetter(input[pos]))
                    {
                        for (int i = pos + 1; i < input.Length &&
                            (Char.IsLetter(input[i]) && input[pos] != ','); i++)
                            s += input[i];
                    }                     
                }
                if (s == "mmax")
                {
                    flag_max = true;
                    int col = 0;
                    yield return "mmax";
                    List<string> num = new List<string>();
                    pos += s.Length;
                    while (input[pos] == ' ') pos++;
                    s = string.Empty + input[pos];
                    yield return s;
                    pos++;
                    while (flag_max && pos < input.Length)
                    {
                        while (input[pos] == ' ') pos++;
                        if (Char.IsDigit(input[pos]))
                        {
                            s = string.Empty;
                            for (int i = pos; i < input.Length && (Char.IsDigit(input[i])); i++)
                                s += input[i];
                            num.Add(s);
                            pos += s.Length;
                            col++;
                        }
                        else if (input[pos] == ',')
                        {
                            pos++;
                        }
                        else if (input[pos] == ')')
                        {
                            
                            for (int i = 0; i < num.Count; i++)
                                yield return num[i];
                            yield return col.ToString();
                            yield return ")";
                            pos++;
                            flag_max = false;
                        }
                    }

                }
                else if (s == "mmin")
                {
                    flag_min = true;
                    int col = 0;
                    yield return "mmin";
                    List<string> num = new List<string>();
                    pos += s.Length;
                    while (input[pos] == ' ') pos++;
                    s = string.Empty + input[pos];
                    yield return s;
                    pos++;
                    while (flag_min && pos < input.Length)
                    {
                        while (input[pos] == ' ') pos++;
                        if (Char.IsDigit(input[pos]))
                        {
                            s = string.Empty;
                            for (int i = pos; i < input.Length && (Char.IsDigit(input[i])); i++)
                                s += input[i];
                            num.Add(s);
                            pos += s.Length;
                            col++;
                        }
                        else if (input[pos] == ',')
                        {
                            pos++;
                        }
                        else if (input[pos] == ')')
                        {
                            
                            for (int i = 0; i < num.Count; i++)
                                yield return num[i];
                            yield return col.ToString();
                            yield return ")";
                            pos++;
                            flag_min = false;
                        }
                    }
                }
                else if (s != ",")
                {
                    yield return s;
                    pos += s.Length;
                }
                else pos++;
            }
        }
        private byte GetPriority(string s)
        {
            switch (s)
            {
                case "(":
                case ")":
                    return 0;
                case "+":
                case "-":
                    return 1;
                case "*":
                case "/":
                    return 2;
                case "^":
                    return 3;
                case "inc":
                case "dec":
                case "min":
                case "max":
                case "mmax":
                case "mmin":
                    return 4;
                default:
                    return 5;
            }
        }
        public string[] ConvertToPostfixNotation(string input)
        {
            List<string> outputSeparated = new List<string>();
            Stack<string> stack = new Stack<string>();

            foreach (string c in Separate(input))
            {
                if (operators.Contains(c))
                {
                    if (stack.Count > 0 && !c.Equals("("))
                    {
                        if (c.Equals(")"))
                        {
                            string s = stack.Pop();
                            while (s != "(")
                            {
                                outputSeparated.Add(s);
                                s = stack.Pop();
                            }
                        }
                        else if (GetPriority(c) > GetPriority(stack.Peek()))
                            stack.Push(c);
                        else
                        {
                            while (stack.Count > 0 && GetPriority(c) <= GetPriority(stack.Peek()))
                                outputSeparated.Add(stack.Pop());
                            stack.Push(c);
                        }
                    }
                    else
                        stack.Push(c);
                }
                else
                    outputSeparated.Add(c);
            }
            if (stack.Count > 0)
                foreach (string c in stack)
                    outputSeparated.Add(c);

            return outputSeparated.ToArray();
        }
        public int result(string input)
        {
            Stack<string> stack = new Stack<string>();
            
            Queue<string> queue = new Queue<string>(ConvertToPostfixNotation(input));
            
            string str = queue.Dequeue();
            
            while (queue.Count >= 0)
            {
                if (!operators.Contains(str))
                { 
                    stack.Push(str);
                    if (queue.Count > 0)
                        str = queue.Dequeue();
                    else break;
                }
                else
                {
                    int summ = 0;
                    try
                    {

                        switch (str)
                        {

                            case "+":
                                {
                                    int a = Convert.ToInt32(stack.Pop());
                                    int b = Convert.ToInt32(stack.Pop());
                                    summ = a + b;
                                    break;
                                }
                            case "-":
                                {
                                    int a = Convert.ToInt32(stack.Pop());
                                    int b = Convert.ToInt32(stack.Pop());
                                    summ = b - a;
                                    break;
                                }
                            case "*":
                                {
                                    int a = Convert.ToInt32(stack.Pop());
                                    int b = Convert.ToInt32(stack.Pop());
                                    summ = b * a;
                                    break;
                                }
                            case "/":
                                {
                                    int a = Convert.ToInt32(stack.Pop());
                                    int b = Convert.ToInt32(stack.Pop());
                                    summ = b / a;
                                    break;
                                }
                            case "^":
                                {
                                    int a = Convert.ToInt32(stack.Pop());
                                    int b = Convert.ToInt32(stack.Pop());
                                    summ = Convert.ToInt32(Math.Pow(Convert.ToDouble(b), Convert.ToDouble(a)));
                                    break;
                                }
                            case "min":
                                {
                                    int a = Convert.ToInt32(stack.Pop());
                                    int b = Convert.ToInt32(stack.Pop());
                                    summ = Math.Min(a, b);
                                    break;
                                }
                            case "max":
                                {
                                    int a = Convert.ToInt32(stack.Pop());
                                    int b = Convert.ToInt32(stack.Pop());
                                    summ = Math.Max(a, b);
                                    break;
                                }
                            case "mmax":
                                {
                                    int k = Convert.ToInt32(stack.Pop());
                                    int b = Convert.ToInt32(stack.Pop());
                                    for (int i = 0; i < k-1; i++)
                                    {
                                        int a = Convert.ToInt32(stack.Pop());
                                        b = Math.Max(b, a);
                                    }
                                    summ = b;
                                    break;
                                }
                            case "mmin":
                                {
                                    int k = Convert.ToInt32(stack.Pop());
                                    int b = Convert.ToInt32(stack.Pop());
                                    for (int i = 0; i < k-1; i++)
                                    {
                                        int a = Convert.ToInt32(stack.Pop());
                                        b = Math.Min(b, a);
                                  
                                    }
                                    summ = b;
                                    break;
                                }
                            case "inc":
                                {
                                    int a = Convert.ToInt32(stack.Pop());
                                    summ = a + 1;
                                    break;
                                }
                            case "dec":
                                {
                                    int a = Convert.ToInt32(stack.Pop());
                                    summ = a - 1;
                                    break;
                                }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    stack.Push(summ.ToString());
                    if (queue.Count > 0)
                        str = queue.Dequeue();
                    else
                        break;
                }

            }
            string ss = stack.Pop();
             
            return Convert.ToInt32(ss);
        }
    
    }
}
