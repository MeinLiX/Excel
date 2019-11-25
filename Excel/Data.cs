using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    enum Mode { Expression, Value
    }

    class Data
    {
        Parser p = new Parser();
        DataGridView dataGridView;
        private int columns = 8;
        private int rows = 8;
        SaveFileDialog saveFileDialog;
        public static List<List<Cell>> cells = new List<List<Cell>>();

        public Data(DataGridView _dataGridView, SaveFileDialog _saveFileDialog)
        {
            dataGridView = _dataGridView;
            cells.Clear();           
            saveFileDialog = _saveFileDialog;
            for (int i = 0; i < rows; i++)
            {
                cells.Add(new List<Cell>());
                for (int j = 0; j < columns; j++)
                {
                    cells[i].Add(new Cell() { RowNumber = i + 1, Columnname = Based26Sys.To26Sys(j) });
                }
            }
        }

        public Parser Parser
        {
            get => default(Parser);
            set
            {
            }
        }

        public Cell Cell
        {
            get => default(Cell);
            set
            {
            }
        }


        public void AddRow()
        {
            cells.Add(new List<Cell>());
            for (int j = 0; j < columns; j++)
            {
                cells[cells.Count - 1].Add(new Cell() { RowNumber = rows + 1, Columnname = Based26Sys.To26Sys(j)});
            }            
            rows++;        
            dataGridView.Rows.Add();
            dataGridView.Rows[dataGridView.Rows.Count - 1].HeaderCell.Value = (dataGridView.Rows.Count).ToString();
            for (int q = 0; q < dataGridView.RowCount; q++)
            {
                for (int p = 0; p < dataGridView.ColumnCount; p++)
                {
                    if (cells[q][p].Expression != null)
                    {
                        cellReload(q, p);
                    }
                }
            }
        }
        public void AddColumn()
        {
            columns++;
            dataGridView.Columns.Add(columns.ToString(), Based26Sys.To26Sys(columns - 1));          
            for (int i = 0; i < rows; i++)
            {
                cells[i].Add(new Cell() { RowNumber = i + 1, Columnname = Based26Sys.To26Sys(columns) });
            }
            for (int q = 0; q < dataGridView.RowCount; q++)
            {
                for (int p = 0; p < dataGridView.ColumnCount; p++)
                {
                    if (cells[q][p].Expression != null)
                    {
                        cellReload(q, p);
                    }
                }
            }
        }
        public void RemoveRow()
        {
            if (rows > 1)
            {
                dataGridView.Rows.RemoveAt(rows - 1);
                cells.RemoveAt(rows - 1);

                for (int i = 0; i < rows - 1; i++)
                {
                    for (int j = 0; j < columns; j++)
                    {
                        if (cells[i][j].References.Where(a => a.RowNumber == rows).Count() != 0)
                        {
                            ChangeData(cells[i][j].Expression, i, j);
                        }
                    }
                }
                rows--;
                for (int q = dataGridView.RowCount-1; q >= 0; q--)
                {
                    for (int p = dataGridView.ColumnCount-1; p >= 0; p--)
                    {
                        if (cells[q][p].Expression != null)
                        {
                            cellReload(q, p);
                        }
                    }
                }
            }
        }

        public void RemoveColumn()
        {
            if (dataGridView.ColumnCount > 1)
            {
                dataGridView.Columns.RemoveAt(columns - 1);
                for (int i = 0; i < rows; i++) 
                {
                    cells[i].RemoveAt(columns - 1);
                }
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < columns-1; j++)
                    {
                        if (cells[i][j].References.Where(a => Based26Sys.From26Sys(a.Columnname) == (columns - 1)).Count() != 0)
                        {
                            ChangeData(cells[i][j].Expression, i, j);
                        }
                    }
                }
                columns--;
                for (int q = dataGridView.RowCount - 1; q >= 0; q--)
                {
                    for (int p = dataGridView.ColumnCount - 1; p >= 0; p--)
                    {
                        if (cells[q][p].Expression != null)
                        {
                            cellReload(q, p);
                        }
                    }
                }
            }
        }
        public void ChangeData(string expression, int row, int col)
        {
            try
            {
                cells[row][col].Expression = expression;
                double res = p.Evaluate(expression, cells[row][col]);
                cells[row][col].Value = res;
                
                cells[row][col].Error = null;

                RecalcReferenceCell(cells[row][col]);
            }
            catch (ParserException ex)
            {
                cells[ex.row][ex.col].Error = ex.Message;
            }
        }
        public void cellReload( int row, int col)
        {
            try
            {
                double res = p.Evaluate(cells[row][col].Expression, cells[row][col]);
                cells[row][col].Value = res;
                cells[row][col].Error = null;
                RecalcReferenceCell(cells[row][col]);
            }
            catch (ParserException ex)
            {
                cells[ex.row][ex.col].Error = ex.Message;
            }
        }
        public void RecalcReferenceCell(Cell cell)
        {
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    if (cells[i][j].Expression != null)
                    {
                        for (int k = 0; k < cells[i][j].References.Count; k++)
                        {
                            if (cells[i][j].References[k].RowNumber == cell.RowNumber && cells[i][j].References[k].Columnname == cell.Columnname)
                            {
                                cells[i][j].Value = p.Evaluate(cells[i][j].Expression, cells[i][j]);
                                cells[i][j].Error = null;
                                dataGridView.Rows[i].Cells[j].Value = cells[i][j].Value.ToString();
                                RecalcReferenceCell(cells[i][j]);
                            }
                        }
                    }
                }
            }
        }
        public void FillData(Mode mode)
        {
            dataGridView.Rows.Clear();
            dataGridView.Columns.Clear();
            for (int i = 0; i < columns; i++) 
            {
                dataGridView.Columns.Add(Based26Sys.To26Sys(i).ToString(), Based26Sys.To26Sys(i).ToString());
            }
            dataGridView.Rows.Add(rows);          
            for (int i = 0; i < rows; ++i)
            {
                dataGridView.Rows[i].HeaderCell.Value = (i + 1).ToString();
                for (int j = 0; j < columns; j++)
                    {
                    if (cells[i][j].Expression != null)
                    {
                        if (cells[i][j].Error != null)
                        { 
                            dataGridView.Rows[i].Cells[j].Value = cells[i][j].Error.ToString();
                        }
                        else
                            dataGridView.Rows[i].Cells[j].Value = mode == Mode.Expression ? cells[i][j].Expression.ToString() : cells[i][j].Value.ToString();
                    }
                }
            }
        }
        public string SaveDataGridView()
        {
            if (saveFileDialog.ShowDialog() == DialogResult.Cancel)
                return null;
            string fileName = saveFileDialog.FileName;
            FileStream aFile = new FileStream(fileName, FileMode.Create);
            StreamWriter sw = new StreamWriter(aFile);
            aFile.Seek(0, SeekOrigin.End);
            sw.WriteLine(columns);
            sw.WriteLine(rows);     
            try
            {
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < columns; j++)
                        if (cells[i][j].Expression != null)
                        {
                            sw.WriteLine(i);
                            sw.WriteLine(j);
                            sw.WriteLine(cells[i][j].Expression);
                            sw.WriteLine(cells[i][j].Value);
                            RecalcReferenceCell(cells[i][j]);
                            if (cells[i][j].Error == null)
                                sw.WriteLine();
                            else sw.WriteLine(cells[i][j].Error);
                        }
                }
                sw.Close();
                MessageBox.Show("Файл збережено", "Excel");
                return fileName;
            }
            catch
            {
                MessageBox.Show("Помилка збереження", "Excel");
                return null;
            }       
        }
       
        public string OpendDataGridView(OpenFileDialog _openFileDialog)
        {
            try
            {
                OpenFileDialog openFileDialog = _openFileDialog;
                DialogResult dialogResult = MessageBox.Show("Зберегти поточну таблицю?", "Excel", MessageBoxButtons.YesNoCancel);
                if (dialogResult == DialogResult.Yes)
                {
                    SaveDataGridView();
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    return null;
                }
                if (openFileDialog.ShowDialog() == DialogResult.Cancel)
                    return null;
                string fileName = openFileDialog.FileName;
                FileStream aFile = new FileStream(fileName, FileMode.Open);
                StreamReader sw = new StreamReader(aFile);
                int totalColumns = Convert.ToInt32(sw.ReadLine());
                int totalRows = Convert.ToInt32(sw.ReadLine());
                while (true)
                    if (totalColumns > columns)
                        AddColumn();
                    else if (totalColumns < columns)
                        RemoveColumn();
                    else break;
                while (true)
                    if (totalRows > rows)
                        AddRow();
                    else if (totalRows < rows)
                        RemoveRow();
                    else break;
                dataGridView.ColumnCount = totalColumns;
                dataGridView.RowCount = totalRows;
                while (!sw.EndOfStream)
                {
                    int i = Convert.ToInt32(sw.ReadLine());
                    int j = Convert.ToInt32(sw.ReadLine());
                    cells[i][j].Expression = sw.ReadLine();
                    cells[i][j].Value = Convert.ToDouble(sw.ReadLine());
                    string error = sw.ReadLine();
                    if (!string.IsNullOrEmpty(error))
                    {
                        cells[i][j].Error = error;
                    }

                }
                for (int q = 0; q < dataGridView.RowCount; q++)
                {
                    for (int p = 0; p < dataGridView.ColumnCount; p++)
                    {
                        if (cells[q][p].Expression != null)
                        {
                            cellReload(q, p);
                        }
                    }
                }

                sw.Close();
                return fileName;
            }
            catch
            {
                MessageBox.Show("Помилка відкриття файлу", "Excel");
                while (true)
                    if (columns > 1) 
                        RemoveColumn();
                    else break;
                while (true)
                    if (rows > 1) 
                        RemoveRow();
                    else break;
                
                while (true)
                    if (rows < 15)
                        AddRow();
                    else break;
                while (true)
                    if (columns < 15)
                        AddColumn();
                    else break;
                

                cells[0][0].Value = 0;
                cells[0][0].Expression = "";             
                return null;               
            }
        }
    }
}

