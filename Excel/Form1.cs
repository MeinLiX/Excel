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

namespace Excel
{
    public partial class Form : System.Windows.Forms.Form
    {
        private Parser p = new Parser();
        Data d;


        public Form()
        {
            InitializeComponent();
            d = new Data(dataGridView, saveFileDialog1);
            this.Load += Form1_Load;
            this.dataGridView.CellParsing += new DataGridViewCellParsingEventHandler(dataGridView_CellParsing);
            this.Closing += new System.ComponentModel.CancelEventHandler(this.FormClosingEventHandler);
            dataGridView.CellBeginEdit += DataGridViewCellBeginEditEventHandler;
            dataGridView.CellEndEdit += DataGridViewCellEndEditEventHandler;
            dataGridView.CellStateChanged += DataGridViewCellStateChangedEventHandler;
        }

        internal Data Data
        {
            get => default(Data);
            set
            {
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            d.FillData(Mode.Value);
        }

        private void dataGridView_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            if (!string.IsNullOrEmpty(e.Value.ToString()))
            {
                d.ChangeData(e.Value.ToString(), e.RowIndex, e.ColumnIndex);

            }
        }
        private void DataGridViewCellStateChangedEventHandler(object sender, DataGridViewCellStateChangedEventArgs e) //+
        {
            try
            {
                if (e.StateChanged.ToString() == "Selected")
                {
                    if (Data.cells[e.Cell.RowIndex][e.Cell.ColumnIndex].Expression != null)
                    {
                        textBox1.Text = Data.cells[e.Cell.RowIndex][e.Cell.ColumnIndex].Expression;
                    }
                    else
                    {
                        textBox1.Text = "";
                    }
                }
            }
            catch
            {

            }
        }
        private void FormClosingEventHandler(object sender, CancelEventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Зберегти таблицю перед завершенням роботи?", "Excel", MessageBoxButtons.YesNoCancel);
            if (dialogResult == DialogResult.Yes)
            {
                d.SaveDataGridView();
                e.Cancel = true;
                return;
            }
            else if (dialogResult == DialogResult.Cancel)
            {
                e.Cancel = true;
                return;
            }
        }
        private void DataGridViewCellBeginEditEventHandler(object sender, DataGridViewCellCancelEventArgs e)
        {
            dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Data.cells[e.RowIndex][e.ColumnIndex].Expression;
        }

        private void DataGridViewCellEndEditEventHandler(object sender, DataGridViewCellEventArgs e)
        {
            if (Data.cells[e.RowIndex][e.ColumnIndex].Expression != null)
                if (!String.IsNullOrEmpty(Data.cells[e.RowIndex][e.ColumnIndex].Error))
                    dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Data.cells[e.RowIndex][e.ColumnIndex].Error;
                else
                    dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Data.cells[e.RowIndex][e.ColumnIndex].Value.ToString();
        }


        private void Tb_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text = ((TextBox)sender).Text;
        }
        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string help = "Лабораторна робота 2\nК-27 \nГрогуль Юрій \nВаріант 10\nБінарні ; \nунарні операції;\n^ (піднесення у степінь);\nmax(x1. x2), mіn(x1. x2);";
            MessageBox.Show(help, "Excel");
        }


        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            d = new Data(dataGridView, saveFileDialog1);
            string name = d.OpendDataGridView(openFileDialog1);
            d.FillData(Mode.Value);

            if (name != null)
                this.Text = "Excel@Hrohul Yurii " + name;
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {

            string name = d.SaveDataGridView();
            if (name != null)
                this.Text = "Excel@Hrohul Yurii " + name;
        }


        private void buttonAdd_Click(object sender, EventArgs e)
        {
            if (radioColumn.Checked == true)
            {
                d.AddColumn();
                d.FillData(Mode.Value);
            }
            else if (radioRow.Checked == true)
            {
                d.AddRow();
                d.FillData(Mode.Value);
            }
        }

        private void buttonRemove_Click(object sender, EventArgs e)
        {
            if (radioColumn.Checked == true)
            {
                DialogResult dialogResult = MessageBox.Show("Дійсно видалити стовпчик?", "Excel", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    d.RemoveColumn();
                }
                d.FillData(Mode.Value);
            }
            else if (radioRow.Checked == true)
            {
                DialogResult dialogResult = MessageBox.Show("Дійсно видалити рядок?", "Excel", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    d.RemoveRow();
                }
                d.FillData(Mode.Value);
            }
        }
    }
}
