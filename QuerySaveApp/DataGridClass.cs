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
        public void DisableTableSorting(DataGridView datagrid)
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
    }
}
