using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace laba2
{

    public partial class Form1 : Form
    {
        string colonki = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string numbers = "1234567890";
        int currRow, currCol;
        public Dictionary<string, MyCell> dictionary = new Dictionary<string, MyCell>(); // ячейка
        public Dictionary<string, string> dictionary2 = new Dictionary<string, string>(); //выражение в ячейке
        public Dictionary<string, string> dictionary3 = new Dictionary<string, string>(); //выражение с подставленными вместо ссылок 

        public Form1()
        {
            InitializeComponent();
            DataGridViewColumn A = new DataGridViewColumn();
            DataGridViewColumn B = new DataGridViewColumn();
            DataGridViewColumn C = new DataGridViewColumn();
            DataGridViewColumn D = new DataGridViewColumn();
            DataGridViewColumn E = new DataGridViewColumn();
            DataGridViewColumn F = new DataGridViewColumn();
            DataGridViewColumn G = new DataGridViewColumn();
            DataGridViewColumn H = new DataGridViewColumn();
            DataGridViewColumn I = new DataGridViewColumn();
            DataGridViewColumn J = new DataGridViewColumn();
            DataGridViewColumn K = new DataGridViewColumn();
            DataGridViewColumn L = new DataGridViewColumn();
            DataGridViewColumn M = new DataGridViewColumn();
            DataGridViewColumn N = new DataGridViewColumn();
            DataGridViewColumn O = new DataGridViewColumn();
            DataGridViewColumn P = new DataGridViewColumn();

            MyCell cellA = new MyCell(); A.CellTemplate = cellA;
            MyCell cellB = new MyCell(); B.CellTemplate = cellB;
            MyCell cellC = new MyCell(); C.CellTemplate = cellC;
            MyCell cellD = new MyCell(); D.CellTemplate = cellD;
            MyCell cellE = new MyCell(); E.CellTemplate = cellE;
            MyCell cellF = new MyCell(); F.CellTemplate = cellF;
            MyCell cellG = new MyCell(); G.CellTemplate = cellG;
            MyCell cellH = new MyCell(); H.CellTemplate = cellH;
            MyCell cellI = new MyCell(); I.CellTemplate = cellI;
            MyCell cellJ = new MyCell(); J.CellTemplate = cellJ;
            MyCell cellK = new MyCell(); K.CellTemplate = cellK;
            MyCell cellL = new MyCell(); L.CellTemplate = cellL;
            MyCell cellM = new MyCell(); M.CellTemplate = cellM;
            MyCell cellN = new MyCell(); N.CellTemplate = cellN;
            MyCell cellO = new MyCell(); O.CellTemplate = cellO;
            MyCell cellP = new MyCell(); P.CellTemplate = cellP;

            A.HeaderText = "A"; A.Name = "A";
            B.HeaderText = "B"; B.Name = "B";
            C.HeaderText = "C"; C.Name = "C";
            D.HeaderText = "D"; D.Name = "D";
            E.HeaderText = "E"; E.Name = "E";
            F.HeaderText = "F"; F.Name = "F";
            G.HeaderText = "G"; G.Name = "G";
            H.HeaderText = "H"; H.Name = "H";
            I.HeaderText = "I"; I.Name = "I";
            J.HeaderText = "J"; J.Name = "J";
            K.HeaderText = "K"; K.Name = "K";
            L.HeaderText = "L"; L.Name = "L";
            M.HeaderText = "M"; M.Name = "M";
            N.HeaderText = "N"; N.Name = "N";
            O.HeaderText = "O"; O.Name = "O";
            P.HeaderText = "P"; P.Name = "P";

            dataGridView.Columns.Add(A);
            dataGridView.Columns.Add(B);
            dataGridView.Columns.Add(C);
            dataGridView.Columns.Add(D);
            dataGridView.Columns.Add(E);
            dataGridView.Columns.Add(F);
            dataGridView.Columns.Add(G);
            dataGridView.Columns.Add(H);
            dataGridView.Columns.Add(I);
            dataGridView.Columns.Add(J);
            dataGridView.Columns.Add(K);
            dataGridView.Columns.Add(L);
            dataGridView.Columns.Add(M);
            dataGridView.Columns.Add(N);
            dataGridView.Columns.Add(O);
            dataGridView.Columns.Add(P);

            DataGridViewRow row = new DataGridViewRow();
            dataGridView.Rows.Add(row); dataGridView.Rows.Add(); dataGridView.Rows.Add();
            dataGridView.Rows.Add(); dataGridView.Rows.Add(); dataGridView.Rows.Add();
            dataGridView.Rows.Add(); dataGridView.Rows.Add(); dataGridView.Rows.Add();
            dataGridView.Rows.Add(); dataGridView.Rows.Add(); dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            dataGridView.Rows.Add();
            SetRowNum(dataGridView);

            for (int i = 0; i < 32; i++)
            {
                for (int j = 0; j < 32; j++)
                {
                    int k = j + 65;
                    string cell_name = (char)k + (i + 1).ToString();
                    //MessageBox.Show(cell_name);
                    MyCell cell = new MyCell();
                    cell.Value = "0";
                    cell.Exp = "0";
                    cell.Depends.Add("");
                    dictionary.Add(cell_name, cell);

                }
            }
        }
        public void SetRowNum(DataGridView dataGrid)
        {
            foreach (DataGridViewRow row in dataGrid.Rows)
            {
                row.HeaderCell.Value = String.Format("{0}", row.Index + 1);
            }
        }
        private void DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DataGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            try
            {
                currRow = dataGridView.CurrentCell.RowIndex;
                currCol = dataGridView.CurrentCell.ColumnIndex;
                string cell_name = "";
                if (currCol < 26)
                    cell_name = (char)(currCol + 65) + (currRow + 1).ToString();
                else
                {
                    cell_name = ((char)(currCol / 25 + 64)).ToString();
                    cell_name += ((char)(currCol - currCol / 25 * 25 + 64)).ToString();
                    cell_name += (currRow + 1).ToString();
                }
                dataGridView[currCol, currRow].Value = dictionary[cell_name].Exp;
                textBox1.Text = dictionary[cell_name].Exp;
            }
            catch { }
        }
        bool flag;
        private void DataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                flag = true;
                currRow = dataGridView.CurrentCell.RowIndex;
                currCol = dataGridView.CurrentCell.ColumnIndex;
                string cell_name = "";
                if (currCol < 26)
                    cell_name = (char)(currCol + 65) + (currRow + 1).ToString();
                else
                {
                    cell_name = ((char)(currCol / 25 + 64)).ToString();
                    cell_name += ((char)(currCol - currCol / 25 * 25 + 64)).ToString();
                    cell_name += (currRow + 1).ToString();
                }
                //MessageBox.Show(((char)(currCol - 25 +  64)).ToString());
                string str = dataGridView[currCol, currRow].Value.ToString();

                dictionary2[cell_name] = str;
                dictionary[cell_name].Exp = str;

                string st = AddressAnalizator(str, cell_name);

                dictionary3[cell_name] = st;
                textBox1.Text = dictionary[cell_name].Exp;
                //MessageBox.Show(st);
                if (IsRecur(cell_name))
                {

                    MessageBox.Show("recur");
                    dataGridView[currCol, currRow].Value = "0";
                    dictionary[cell_name].Value = "0";
                    dictionary[cell_name].Exp = "0";

                    dictionary2[cell_name] = null;
                    dictionary3[cell_name] = null;
                    Update(cell_name, cell_name);
                    dictionary[cell_name].Depends.Clear();
                    dictionary[cell_name].Depends_cells.Clear();
                }
                else
                {

                    Parser pars = new Parser();
                    int res = pars.result(st);
                    string result = res.ToString();
                    //MessageBox.Show(cell_name);
                    if (pars.str_error != "")
                    {
                        MessageBox.Show(pars.str_error);

                        dataGridView[currCol, currRow].Value = "0";

                        dictionary[cell_name].Value = "0";
                        dictionary[cell_name].Exp = "0";
                        dictionary2[cell_name] = null;
                        dictionary3[cell_name] = null;
                        pars.str_error = "";
                        Update(cell_name, cell_name);
                        //MessageBox.Show(cell_name);
                    }
                    else
                    {
                        dictionary[cell_name].Value = result;

                        dataGridView[currCol, currRow].Value = dictionary[cell_name].Value;
                        textBox1.Text = dictionary3[cell_name];
                        Update(cell_name, cell_name);
                    }
                }
                textBox1.Text = "";
                RefreshCells();
            }
            catch { }
        }
        public void Update(string cell_name, string nameSt)
        {

            for (int i = 0; i < dictionary[cell_name].Depends_cells.Count(); i++)
            {
                string name = dictionary[cell_name].Depends_cells[i];
                //MessageBox.Show(name);
                if (name == nameSt)
                {
                    return;
                }
                string str = dictionary[name].Exp;
                //MessageBox.Show(str);
                string st = AddressAnalizator(str, name);
                dictionary3[name] = st;
                Parser pars = new Parser();
                int res = pars.result(st);
                string result = res.ToString();
                if (pars.str_error != "")
                {
                    MessageBox.Show(pars.str_error);
                    dataGridView[currCol, currRow].Value = "0";
                    dictionary[name].Value = "0";
                    dictionary[name].Exp = "0";
                    dictionary2[name] = null;
                    dictionary3[name] = null;
                    pars.str_error = "";
                    Update(name, nameSt);
                }
                else
                {
                    dictionary[name].Value = result;
                    int coll = 0, row = 0;
                    if (colonki.Contains(name[0]))
                    {
                        if (colonki.Contains(name[1]))
                        {
                            coll = Convert.ToInt32(name[0]) - 65 + Convert.ToInt32(name[1]) - 65 + 26;
                            string a = "";
                            for (int k = 2; k < name.Length; k++)
                            {
                                a += name[k].ToString();
                            }
                            row = Convert.ToInt32(a) - 1;
                        }
                        else
                        {
                            coll = Convert.ToInt32(name[0]) - 65;
                            string a = "";
                            for (int k = 1; k < name.Length; k++)
                            {
                                a += name[k].ToString();
                            }
                            row = Convert.ToInt32(a) - 1;
                        }
                    }
                    dataGridView[coll, row].Value = dictionary[name].Value;
                }
                Update(name, nameSt);
            }
        }
        public string AddressAnalizator(string str, string cell_name)
        {
            string st = "";
            char[] delim = { ' ' };
            List<string> lex = new List<string>(str.Split(delim));
            for (int i = 0; i < lex.Count; i++)
            {

                if (colonki.Contains(lex[i][0]))
                {

                    if (dictionary.ContainsKey(lex[i]))
                    {
                        // MessageBox.Show(dictionary[lex[i]].Value);
                        if (!dictionary[lex[i]].Depends_cells.Contains(cell_name))
                            dictionary[lex[i]].Depends_cells.Add(cell_name);
                        lex[i] = dictionary[lex[i]].Value;
                        if (!dictionary[cell_name].Depends.Contains(lex[i]))
                            dictionary[cell_name].Depends.Add(lex[i]);



                        // MessageBox.Show(lex[i]);

                    }
                    else
                    {
                        MessageBox.Show("Incorrect adress");
                        dataGridView.CurrentCell.Value = "0";
                        lex[i] = "0";
                        return "0";
                    }

                }

            }
            for (int i = 0; i < lex.Count; i++)
            {
                st += lex[i];
            }
            str = st;
            return str;
        }
        void RefreshCells()
        {
            dataGridView.Update();
            dataGridView.Refresh();
        }
        bool dop_rec(string celll_name, string celll_name1)
        {
            // MessageBox.Show(celll_name);
            if (dictionary[celll_name1].Exp.Contains(celll_name))
                return true;
            else
            {
                char[] delim = { ' ' };
                string[] dep = (dictionary[celll_name1].Exp.Split(delim));
                for (int i = 0; i < dep.Length; i++)
                {
                    //string cellname = "";
                    if (colonki.Contains(dep[i][0]))
                    {
                        if (dictionary.ContainsKey(dep[i]))
                        {
                            bool w = dop_rec(celll_name, dep[i]);
                            if (w == true)
                            {
                                return true;
                            }
                        }
                    }

                }
            }
            return false;
        }
        public bool IsRecur(string cell_name)
        {
            return dop_rec(cell_name, cell_name);
        }

        private void AddRow()
        {
            DataGridViewRow row = new DataGridViewRow();
            dataGridView.Rows.Add(row);
            SetRowNum(dataGridView);
            for (int col = 0; col < dataGridView.ColumnCount; col++)
            {
                string cell_name = "";
                if (col < 26)
                    cell_name = (char)(col + 65) + (dataGridView.RowCount + 1).ToString();
                else
                {
                    cell_name = ((char)(currCol / 25 + 64)).ToString();
                    cell_name += ((char)(currCol - currCol / 25 * 25 + 64)).ToString();
                    cell_name += (dataGridView.RowCount + 1).ToString();
                }
                MyCell cell = new MyCell();
                cell.Value = "0";
                cell.Exp = "0";
                cell.Depends.Add("");
                if (!dictionary.ContainsKey(cell_name))
                    dictionary.Add(cell_name, cell);
                if (!dictionary2.ContainsKey(cell_name))
                    dictionary2.Add(cell_name, null);
                if (!dictionary3.ContainsKey(cell_name))
                    dictionary3.Add(cell_name, null);
                // MessageBox.Show(cell_name);
            }

            RefreshCells();
        }

        private void AddColumn()
        {
            DataGridViewColumn column = new DataGridViewColumn();
            DataGridViewCell cell = new DataGridViewTextBoxCell();
            column.CellTemplate = cell;
            int k = dataGridView.ColumnCount - 1;
            string n = dataGridView.Columns[k].Name;
            string cc = "";
            if (n.Length == 1)
            {
                if (n == "Z")
                {
                    cc = "AA";

                }
                else
                {
                    int j = colonki.IndexOf(n[0]);
                    cc = colonki[j + 1].ToString();
                }
            }
            else
            {
                if (n[1] == 'Z')
                {
                    int j = colonki.IndexOf(n[0]);
                    cc = colonki[j + 1].ToString() + "A";
                }
                else
                {
                    int j = colonki.IndexOf(n[1]);
                    cc = n[0].ToString() + colonki[j + 1].ToString();
                }
            }


            column.HeaderText = cc;
            column.Name = cc;
            dataGridView.Columns.Add(column);
            for (int row = 0; row < dataGridView.RowCount; row++)
            {
                string cell_name = cc + (row + 1).ToString();
                MyCell celll = new MyCell();
                celll.Value = "0";
                celll.Exp = "0";
                celll.Depends.Add("");
                if (!dictionary.ContainsKey(cell_name))
                    dictionary.Add(cell_name, celll);
                if (!dictionary2.ContainsKey(cell_name))
                    dictionary2.Add(cell_name, null);
                if (!dictionary3.ContainsKey(cell_name))
                    dictionary3.Add(cell_name, null);
                // MessageBox.Show(cell_name);
            }
            RefreshCells();
        }

        private void DeleteColumn()
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("Данні кряйнього рядку будуть видалені.\nДісно видалити?", "Delete Column", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    if (dataGridView.Columns.Count > 1)
                    {
                        int col = dataGridView.ColumnCount - 1;
                        dataGridView.Columns.RemoveAt(col);
                        for (int row = 0; row < dataGridView.RowCount; row++)
                        {
                            string cell_name = "";
                            if (col < 26)
                                cell_name = (char)(col + 65) + (row + 1).ToString();
                            else
                            {
                                cell_name = ((char)(currCol / 25 + 64)).ToString();
                                cell_name += ((char)(currCol - currCol / 25 * 25 + 64)).ToString();
                                cell_name += (currRow + 1).ToString();
                            }
                            if (dictionary[cell_name].Depends_cells.Count > 0)
                                MessageBox.Show("Incorrect adress");
                            dictionary[cell_name].Value = "0";
                            dictionary[cell_name].Exp = "0";
                            dictionary2[cell_name] = null;
                            dictionary3[cell_name] = null;
                            Update(cell_name, cell_name);

                            dictionary.Remove(cell_name);
                            dictionary2.Remove(cell_name);
                            dictionary3.Remove(cell_name);
                        }
                        RefreshCells();
                    }
                }
            }
            catch { }
        }

        private void DeleteRow_Click(object sender, EventArgs e)
        {
            if (dataGridView.Rows.Count > 1)
                RowDelete(dataGridView);
        }
        private void RowDelete(DataGridView dataGrid)
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("Данні кряйнього рядку будуть видалені.\nДісно видалити?", "Delete Row", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    int row = dataGrid.RowCount - 2;
                    dataGrid.Rows.RemoveAt(row);
                    for (int col = 0; col < dataGrid.ColumnCount; col++)
                    {
                        string cell_name = "";
                        if (col < 26)
                            cell_name = (char)(col + 65) + (row + 1).ToString();
                        else
                        {
                            cell_name = ((char)(currCol / 25 + 64)).ToString();
                            cell_name += ((char)(currCol - currCol / 25 * 25 + 64)).ToString();
                            cell_name += (currRow + 1).ToString();
                        }

                        if (dictionary[cell_name].Depends_cells.Count > 0)
                            MessageBox.Show("Incorrect adress");
                        dictionary[cell_name].Value = "0";
                        dictionary[cell_name].Exp = "0";
                        dictionary2[cell_name] = null;
                        dictionary3[cell_name] = null;
                        Update(cell_name, cell_name);
                        dictionary.Remove(cell_name);
                        dictionary2.Remove(cell_name);
                        dictionary3.Remove(cell_name);
                    }
                    RefreshCells();
                    SetRowNum(dataGrid);
                }
            }
            catch { }
        }
        private void closeFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFile();
        }
        private void SaveFile()
        {
            Stream mystream;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                flag = false;
                if ((mystream = saveFileDialog1.OpenFile()) != null)
                {
                    StreamWriter sw = new StreamWriter(mystream);
                    sw.WriteLine(dataGridView.RowCount);
                    sw.WriteLine(dataGridView.ColumnCount);
                    for (int i = 0; i < dataGridView.RowCount; i++)
                    {
                        for (int j = 0; j < dataGridView.ColumnCount; j++)
                        {
                            if (dataGridView.Rows[i].Cells[j].Value != null)
                            {
                                string cell_name = "";
                                if (j < 26)
                                    cell_name = (char)(j + 65) + (i + 1).ToString();
                                else
                                {
                                    cell_name = ((char)(j / 25 + 64)).ToString();
                                    cell_name += ((char)(j - j / 25 * 25 + 64)).ToString();
                                    cell_name += (i + 1).ToString();
                                }
                                sw.WriteLine(cell_name);
                                sw.WriteLine(dictionary[cell_name].Exp);
                                sw.WriteLine(dictionary[cell_name].Value);
                                sw.WriteLine(dictionary3[cell_name]);
                            }

                        }
                    }
                    pathh.Text = openFileDialog1.FileName;
                    sw.Close();
                    mystream.Close();
                }
            }
        }
        private void openFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFile();
        }
        private void OpenFile()
        {
            DialogResult dialogResult = MessageBox.Show("Відкритий новий файл?\nВсі данні будуть видалені.", "Open file", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                Stream mystream = null;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    if ((mystream = openFileDialog1.OpenFile()) != null)
                    {
                        using (mystream)
                        {
                            try
                            {
                                StreamReader sr = new StreamReader(mystream);
                                string scr = sr.ReadLine();
                                string scc = sr.ReadLine();
                                int cr = Convert.ToInt32(scr);
                                int cc = Convert.ToInt32(scc);
                                for (int i = 0; i < cr; i++)
                                {
                                    for (int j = 0; j < cc; j++)
                                    {
                                        string name = "";
                                        if (j < 26)
                                            name = (char)(j + 65) + (i + 1).ToString();
                                        else
                                        {
                                            name = ((char)(j / 25 + 64)).ToString();
                                            name += ((char)(j - j / 25 * 25 + 64)).ToString();
                                            name += (i + 1).ToString();
                                        }

                                        if (dictionary[name].Exp != "0")
                                        {
                                            dataGridView[j, i].Value = null;
                                            dictionary[name].Exp = "0";
                                            dictionary2[name] = "0";
                                            dictionary[name].Value = "0";
                                            dictionary3[name] = "0";
                                        }
                                    }
                                }

                                while (dataGridView.Rows.Count < cr)
                                {
                                    DataGridViewRow row = new DataGridViewRow();
                                    dataGridView.Rows.Add(row);
                                }
                                while (dataGridView.Columns.Count < cc)
                                {
                                    DataGridViewColumn column = new DataGridViewColumn();
                                    dataGridView.Columns.Add(column);
                                }

                                while (!sr.EndOfStream)
                                {
                                    string name = sr.ReadLine();

                                    int coll = 0, row = 0;
                                    if (colonki.Contains(name[0]))
                                    {
                                        if (colonki.Contains(name[1]))
                                        {
                                            coll = Convert.ToInt32(name[0]) - 65 + Convert.ToInt32(name[1]) - 65 + 26;
                                            string a = "";
                                            for (int k = 2; k < name.Length; k++)
                                            {
                                                a += name[k].ToString();
                                            }
                                            row = Convert.ToInt32(a) - 1;
                                        }
                                        else
                                        {
                                            coll = Convert.ToInt32(name[0]) - 65;
                                            string a = "";
                                            for (int k = 1; k < name.Length; k++)
                                            {
                                                a += name[k].ToString();
                                            }
                                            row = Convert.ToInt32(a) - 1;
                                        }
                                    }
                                    string exp = sr.ReadLine();
                                    string val = sr.ReadLine();
                                    string d3 = sr.ReadLine();
                                    dataGridView[coll, row].Value = val;
                                    if (!dictionary.ContainsKey(name))
                                    {
                                        MyCell cell = new MyCell();
                                        cell.Value = "0";
                                        cell.Exp = "0";
                                        cell.Depends.Add("");
                                        dictionary.Add(name, cell);
                                        dictionary2.Add(name, "0");
                                        dictionary3.Add(name, "0");
                                    }
                                    dictionary[name].Exp = exp;
                                    dictionary2[name] = exp;
                                    dictionary[name].Value = val;
                                    dictionary3[name] = d3;
                                }
                                sr.Close();
                                pathh.Text = openFileDialog1.FileName;
                            }
                            catch { }
                        }
                    }
                }
            }
        }

        private void DataGridView_MouseClick(object sender, MouseEventArgs e)
        {
            currRow = dataGridView.CurrentCell.RowIndex;
            currCol = dataGridView.CurrentCell.ColumnIndex;
            string cell_name = "";
            if (currCol < 26)
                cell_name = (char)(currCol + 65) + (currRow + 1).ToString();
            else
            {
                cell_name = ((char)(currCol / 25 + 64)).ToString();
                cell_name += ((char)(currCol - currCol / 25 * 25 + 64)).ToString();
                cell_name += (currRow + 1).ToString();
            }
            try
            {
                textBox1.Text = dictionary[cell_name].Exp;
            }
            catch { };
        }
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About();
        }
        private void About()
        {
            MessageBox.Show("К-27\nВаріант 11:\n1)+,-,*,/(бінарні операції) \n4)^ (піднесення у степінь) \n5)іnc, dec \n6)max(x,y),min(x,y)", "Excel");
        }

        private void Add_Click(object sender, EventArgs e)
        {
            if (radioColumn.Checked == true)
            {
                AddColumn();
            }
            if (radioRow.Checked == true)
            {
                AddRow();
            }
        }

        private void Remove_Click(object sender, EventArgs e)
        {
            if (radioColumn.Checked == true)
            {
                DeleteColumn();
            }
            if (radioRow.Checked == true)
            {
                RowDelete(dataGridView);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Закрити програму?\nВсі данні будуть видалені.", "Close file", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No)
                e.Cancel = true;

        }
    }
}
