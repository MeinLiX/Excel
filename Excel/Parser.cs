using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public class ParserException : ApplicationException
    {
        public int row;
        public int col;
        public ParserException(string str, int row, int col) : base(str)
        {
            this.row = row;
            this.col = col;
        }
        public override string ToString()
        { return Message; }

        public void Method()
        {
            throw new System.NotImplementedException();
        }
    }

    public class Parser
    {
        enum Types { NONE, DELIMITER, VARIABLE, NUMBER };
        enum Errors { SYNTAX, UNBALPARENS, NOEXP, DIVBYZERO, RECURCELLS, REFWITHERROR, NOTEXTINCELL };
        string exp;
        int expIdx;
        string token;
        Types tokType;
        double[] vars = new double[26];

        Cell currentCell;
        int currentRowCell;
        int currentColumnCell;

        public Parser()
        {
            for (int i = 0; i < vars.Length; i++)
                vars[i] = 0.0;
        }

        public Cell Cell
        {
            get => default(Cell);
            set
            {
            }
        }

        public ParserException ParserException
        {
            get => default(ParserException);
            set
            {
            }
        }

        public double Evaluate(string expstr, Cell _currentCell)
        {
            currentCell = _currentCell;
            currentColumnCell = Based26Sys.From26Sys(_currentCell.Columnname);
            currentRowCell = _currentCell.RowNumber - 1;
            currentCell.References.Clear();
            double result;
            expstr = EvalExp0(expstr);

            exp = expstr;
            expIdx = 0;
            GetToken();
            if (token == "close")
            {
                SyntaxErr(Errors.RECURCELLS);
                return 0.0;
            }
            if (token == "")
            {
                SyntaxErr(Errors.NOEXP);
                return 0.0;
            }
            EvalExp1(out result);
            if (token != "")
                SyntaxErr(Errors.SYNTAX);

            if (CheckRecurInCells(currentCell))
                SyntaxErr(Errors.RECURCELLS);

            return result;
        }
        bool CheckRecurInCells(Cell cell)
        {
            foreach (var i in cell.References)
            {

                if (i == currentCell)
                {
                    return true;
                }
                if (CheckRecurInCells(i) == true)
                {
                    return true;
                }
            }
            return false;
        }

        private string EvalExp0(string result)//min(.), max(.)
        {
            try
            {
                if (result.Substring(0, 3) == "min" && result.IndexOf(".") != -1)
                {
                    result = result.Remove(0, result.IndexOf("("));
                    result = result.Replace(".", "<");
                }
                else if (result.Substring(0, 3) == "max" && result.IndexOf(".") != -1)
                {
                    result = result.Remove(0, result.IndexOf("("));
                    result = result.Replace(".", ">");
                }
                else if ((result.Substring(0, 3) == "max") || (result.Substring(0, 3) == "min")) MessageBox.Show("Uncorrect Function", "MyMiniExcel ERR!");
            }
            catch (Exception)
            {
            }
            return result;
        }
        void EvalExp1(out double result)
        {
            int varIdx;
            Types ttokType;
            string temptoken;
            if (tokType == Types.VARIABLE)
            {
                temptoken = String.Copy(token);
                ttokType = tokType;
                varIdx = Char.ToUpper(token[0]) - 'A';
                GetToken();
                if (token != "=")
                {
                    PutBack();
                    token = String.Copy(temptoken);
                    tokType = ttokType;
                }
                else
                {
                    GetToken();
                    EvalExp2(out result);
                    vars[varIdx] = result;
                    return;
                }
            }
            EvalExp2(out result);
        }
        void EvalExp2(out double result)
        {
            string op;
            double partialResult;
            EvalExp3(out result);
            while ((op = token) == "+" || op == "-")
            {
                GetToken();
                EvalExp3(out partialResult);
                switch (op)
                {
                    case "-":
                        result = result - partialResult;
                        break;
                    case "+":
                        result = result + partialResult;
                        break;
                }
            }
        }
        void EvalExp3(out double result)
        {
            string op;
            double partialResult = 0.0;
            EvalExp4(out result);
            while ((op = token) == "*" || op == "/" || op == "%" || op == "|")
            {
                GetToken();
                EvalExp4(out partialResult);
                switch (op)
                {
                    case "*":
                        result = result * partialResult;
                        break;
                    case "/":
                        if (partialResult == 0.0)
                            SyntaxErr(Errors.DIVBYZERO);
                        result = result / partialResult;
                        break;
                    case "%":
                        if (partialResult == 0.0)
                            SyntaxErr(Errors.DIVBYZERO);
                        result = (int)result % (int)partialResult;
                        break;
                    case "|":
                        if (partialResult == 0.0)
                            SyntaxErr(Errors.DIVBYZERO);
                        result = (int)result / (int)partialResult;
                        break;
                }
            }
        }

        void EvalExp4(out double result)
        {
            double partialResult, ex;
            int t;
            EvalExp5(out result);
            if (token == "^")
            {
                GetToken();
                EvalExp4(out partialResult);
                ex = result;
                if (partialResult == 0.0)
                {
                    result = 1.0;
                    return;
                }
                for (t = (int)partialResult - 1; t > 0; t--)
                    result = result * (double)ex;
            }
        }
        private void EvalExp5(out double result)
        {
            string op;
            double partialResult;
            EvalExp6(out result);
            while ((op = token) == ">" || op == "<")
            {
                GetToken();
                EvalExp6(out partialResult);
                switch (op)
                {
                    case ">":
                        if (result < partialResult)
                        {
                            result = partialResult;
                        }
                        break;
                    case "<":
                        if (result > partialResult)
                        {
                            result = partialResult;
                        }
                        break;
                }
            }
        }

        void EvalExp6(out double result)
        {
            string op;
            op = "";
            if ((tokType == Types.DELIMITER) && token == "+" || token == "-")
            {
                op = token;
                GetToken();
            }
            EvalExp7(out result);
            if (op == "-") result = -result;
        }

        void EvalExp7(out double result)
        {
            if ((token == "("))
            {
                GetToken();
                EvalExp2(out result);
                if (token != ")")
                {
                    SyntaxErr(Errors.UNBALPARENS);
                }
                GetToken();
            }
            else Atom(out result);
        }
        void Atom(out double result)
        {
            switch (tokType)
            {
                case Types.NUMBER:
                    try
                    {
                        result = Double.Parse(token);
                    }
                    catch (FormatException)
                    {
                        result = 0.0;
                        SyntaxErr(Errors.SYNTAX);
                    }
                    GetToken();
                    return;
                case Types.VARIABLE:
                    result = FindVar(token);
                    GetToken();
                    return;
                default:
                    result = 0.0;
                    SyntaxErr(Errors.SYNTAX);
                    break;
            }
        }

        double FindVar(string vname)
        {
            if (!Char.IsLetter(vname[0]))
            {
                SyntaxErr(Errors.SYNTAX);
                return 0.0;
            }
            return vars[Char.ToUpper(vname[0]) - 'A'];
        }
        void PutBack()
        {
            for (int i = 0; i < token.Length; i++) expIdx--;
        }
        void SyntaxErr(Errors error)
        {
            string[] err ={

            "Синтаксична помилка",
            "Дизбаланс дужок",
            "Вираз відсутній",
            "Ділення на нуль",
            "Рекурсивні посилання",
            "Посилання на клітинку з помилкою",
            "Посилання на неіснуючу клітинку"};
            throw new ParserException(err[(int)error], currentRowCell, currentColumnCell);
        }
        private void GetToken()
        {
            try
            {
                tokType = Types.NONE;
                token = "";
                if (expIdx == exp.Length)
                    return;
                while (expIdx < exp.Length && char.IsWhiteSpace(exp[expIdx]))
                    ++expIdx;
                if (expIdx == exp.Length)
                    return;
                if (IsDelim(exp[expIdx]))
                {
                    token += exp[expIdx];
                    expIdx++;
                    tokType = Types.DELIMITER;
                }
                else if (Char.IsLetter(exp[expIdx]))
                {
                    while (!IsDelim(exp[expIdx]))
                    {
                        token += exp[expIdx];
                        expIdx++;
                        if (expIdx >= exp.Length)
                            break;
                    }
                    int i = 0;
                    string column = null;
                    string row = null;
                    int currentRow;
                    int currentColumn;
                    while (Char.IsLetter(token[i]))
                    {
                        column += token[i];
                        i++;
                    }
                    while (i < token.Length)
                    {
                        row += token[i];
                        i++;
                    }
                    if (!int.TryParse(row, out currentRow))
                    {
                        SyntaxErr(Errors.SYNTAX);
                    }
                    currentColumn = Based26Sys.From26Sys(column);
                    currentRow = Convert.ToInt32(row) - 1;
                    if (currentRow >= Data.cells.Count || currentColumn >= Data.cells[currentRow].Count)
                    {
                        Cell notExistCell = new Cell() { RowNumber = currentRow + 1, Columnname = column };
                        currentCell.References.Add(notExistCell);
                        SyntaxErr(Errors.REFWITHERROR);
                    }
                    try
                    {
                        Cell parsedCell = Data.cells[currentRow][currentColumn];
                        currentCell.References.Add(parsedCell);


                        if (parsedCell.Error == "Рекурсивні посилання")
                        {
                            token = "close";
                            return;
                        }
                        if (!string.IsNullOrEmpty(parsedCell.Error) && !CheckRecurInCells(parsedCell))
                        {
                            SyntaxErr(Errors.NOTEXTINCELL);

                        }

                        token = parsedCell.Value.ToString();
                    }
                    catch
                    {

                    }

                    tokType = Types.NUMBER;
                }
                else if (Char.IsDigit(exp[expIdx]))
                {
                    while (!IsDelim(exp[expIdx]))
                    {
                        token += exp[expIdx];
                        expIdx++;
                        if (expIdx >= exp.Length)
                            break;
                    }

                    tokType = Types.NUMBER;
                }
            }
            catch
            {

            }
        }
        private bool IsDelim(char c)
        {
            if (("+-/*%=()^|><".IndexOf(c) != -1))
                return true;
            return false;
        }
    }
}