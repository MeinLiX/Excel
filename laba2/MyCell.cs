using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace laba2
{
    public class MyCell :  DataGridViewTextBoxCell
    {
        string val;
        string name;
        string exp;
        List<string> dep = new List<string>();
        List<string> dep_n = new List<string>();
        public MyCell()
        {
            name = "A" + " 1";
            val = "";
            exp = "";
        }
        public string Name
        {
            get { return name; }
        }
        public string Value
        {
            get { return val; }
            set { val = value; }
        }
        public string Exp
        {
            get { return exp; }
            set { exp = value; }
        }
        public List<string> Depends
        {
            get { return dep; }
            set { dep = value; }
        }
        public List<string> Depends_cells
        {
            get { return dep_n; }
            set { dep_n = value; }
        }
    }
}
