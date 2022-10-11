using Excel = Microsoft.Office.Interop.Excel;
// excel object library is required to use the above, not the fucking shitty interlop piece of shit.
using System.Reflection;//fuck knows what this one is being used for.
// gotta use system not microsoft or it errors on using a certificate for some reason.
using System.Data.SqlClient;
using System;
using System.Data;


namespace QuerySaveApp
{
    public partial class QuerySave : Form
    {
        public QuerySave()
        {
            InitializeComponent();
            DataGridClass dataGridClass = new DataGridClass();

            //set Datasource and Datamember, both of these are required to return data.
            dataGridView1.DataSource = XMLData.ReturnXMLDataset();
            dataGridView1.DataMember = "Authors";

            ////add a button column and then add buttons to it
            dataGridClass.addButtonColumn(dataGridView1, "Browse save location", "Run", "BrowseButton", "browse");
            dataGridClass.addButtonColumn(dataGridView1, "Run Report", "Run", "ReportButton", "Run");

            dataGridClass.DisableTableSorting(dataGridView1);

        }


  

        private void button2_Click(object sender, EventArgs e)
        {
            //SAVES THE DATASET.

            //what is this (dataset)
            DataSet ds = (DataSet)dataGridView1.DataSource;
            ds.WriteXml(XMLData.getFilePath());
        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

            //richTextBox1.AppendText(dataGridView1[e.ColumnIndex, e.RowIndex].ToString()!);

            // this enables you to access custom generated buttons,  DataGridViewCellEventArgs is passing information about the location of the pressed cell.
            // e.columnindex and e.rowindex can be used to access the specific cell.
            try
            {
                if (dataGridView1[e.ColumnIndex, e.RowIndex] is DataGridViewButtonCell cell)
                {
                   
                 

                    if (e.ColumnIndex == 0)
                    {

                        SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
                        DialogResult result = SaveFileDialog1.ShowDialog(); // Show the dialog.

                        if (result == DialogResult.OK) // Test result.
                        {


                            string file = SaveFileDialog1.FileName;

                            dataGridView1[3, e.RowIndex].Value = file;

                            richTextBox1.AppendText(e.ColumnIndex.ToString());

                            //cell.ColumnIndex[j] ,  = file;

                        }
                    }
                    else if (e.ColumnIndex == 1)
                    {


                        ExcelWrite.xlsWorkBook(dataGridView1[2, e.RowIndex].Value.ToString()!, dataGridView1[3, e.RowIndex].Value.ToString()!);
               
                    }

                }

                else { return; }
            }
            catch (ArgumentOutOfRangeException)
            {
                richTextBox1.AppendText("Argument out of range");
            }

        }
    }
}