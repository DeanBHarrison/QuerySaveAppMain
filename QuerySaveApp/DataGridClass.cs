using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuerySaveApp
{
    internal class DataGridClass
    {
        
        //function to disable column sorting on a given datagridview.
        public  static void DisableTableSorting(DataGridView datagrid)
        {
            //disable sorting for each column
            foreach (DataGridViewColumn column in datagrid.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }


        //add a button column and then add buttons to it
        public void addButtonColumn(DataGridView datagrid, String Columnname, string buttonTxt, string buttonname, string visibleName)
        {
            DataGridViewButtonColumn btnColumn = new DataGridViewButtonColumn();
            datagrid.Columns.Add(btnColumn);
            btnColumn.HeaderText = Columnname;
            //btnColumn.Text = buttonTxt;
            //btnColumn.Name = buttonname;
            btnColumn.UseColumnTextForButtonValue = false;
            btnColumn.DefaultCellStyle = new DataGridViewCellStyle()
            {
                NullValue = visibleName
            };       
        }


        // add a dropdown column
        public void addcombo(DataGridView datagrid)
        {
            DataGridViewComboBoxColumn inputtablecombobox = new DataGridViewComboBoxColumn();
            inputtablecombobox.HeaderText = "field";
            inputtablecombobox.Name = "inputtablecombobox";
            //String combosql = "select field from input_metadata";
            //SqlDataAdapter comboadapter = new SqlDataAdapter(combosql, connection);
            //DataSet ds = new DataSet();
            //comboadapter.Fill(ds);
            //inputtablecombobox.DataSource = ds.Tables[0];
            //inputtablecombobox.DisplayMember = "field";
            //inputtablecombobox.ValueMember = "field";
            datagrid.Columns.Add(inputtablecombobox);   
        }

        public  static void SetColumnsOrder(DataGridView datagrid, params String[] columnNames)
        {
            int columnIndex = 0;
            foreach (var columnName in columnNames)
            {
                datagrid.Columns[columnName].DisplayIndex = columnIndex;
                columnIndex++;
            }
        }

        
        public static void LoadButtons(DataGridView dataGridView)
        {
            for (int i = 0; i < dataGridView.RowCount; i++)
            {
                for (int j = 0; j < dataGridView.ColumnCount; j++)
                {
                    if (dataGridView.Rows[i].Cells[j].Value.ToString() == "Browse" || dataGridView.Rows[i].Cells[j].Value.ToString() == "Run" || dataGridView.Rows[i].Cells[j].Value.ToString() == "Settings")
                    {
                        //richTextBox1.AppendText(dataGridView.Rows[i].Cells[j].Value.ToString());
                        // Here is the trick.
                        var btnCell = new DataGridViewButtonCell();
                        dataGridView.Rows[i].Cells[j] = btnCell;
                    }
                }
            }
        }

        //HERE
        public static void LoadCombocells(DataGridView dataGridView)
        {
            for (int i = 0; i < dataGridView.RowCount; i++)
            {
                for (int j = 0; j < dataGridView.ColumnCount; j++)
                {
                    if (dataGridView.Rows[i].Cells[j].Value.ToString() == "Yes" || dataGridView.Rows[i].Cells[j].Value.ToString() == "No" )
                    {
                        //richTextBox1.AppendText(dataGridView.Rows[i].Cells[j].Value.ToString());
                        // Here is the trick.
                        var combobtnCell = new DataGridViewComboBoxCell();
                        dataGridView.Rows[i].Cells[j] = combobtnCell;
                    }
                }
            }
        }

        public static string CellColumn(DataGridView dataGridView1, string ColumnName, int Index)
        {
            var loadindex = dataGridView1.Columns[ColumnName].Index;
            var loadstring = dataGridView1[loadindex, Index].Value.ToString()!;

            return loadstring;
        }

        public static string SaveLocationReturn(DataGridView dataGridView1, string ColumnName, int Index)
        {
            var loadindex = dataGridView1.Columns[ColumnName].Index;
            var loadstring = dataGridView1[loadindex, Index].Value.ToString()!;

            return loadstring;
        }
    }
}
