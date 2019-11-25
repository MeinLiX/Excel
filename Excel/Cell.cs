using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public class Cell
    {
        public string Expression { get; set;}
        public double Value { get; set; }
        public string Error { get; set; }       
        public int RowNumber;
        
        public string Columnname;
        public List<Cell> References { get; set;} = new List<Cell>();

        public void Method()
        {
            throw new System.NotImplementedException();
        }

        public Based26Sys Based26Sys
        {
            get => default(Based26Sys);
            set
            {
            }
        }
    }
}
