using Excel = Microsoft.Office.Interop.Excel;
// excel object library is required to use the above, not the fucking shitty interlop piece of shit.
using System.Reflection;//fuck knows what this one is being used for.
// gotta use system not microsoft or it errors on using a certificate for some reason.
using System.Data.SqlClient;
using System;
using System.Data;
using GXOMIClassLibrary;


namespace QuerySaveApp
{
    public partial class QuerySave : Form
    {
        public QuerySave()
        {
            InitializeComponent();
            //set Datasource and Datamember, both of these are required to return data.
            dataGridView1.DataSource = XMLData.ReturnXMLDataset();
            dataGridView1.DataMember = "Authors";

            ////add a button column and then add buttons to it
            DataGridClass dataGridClass = new DataGridClass();
            //dataGridClass.addButtonColumn(dataGridView1, "Browse save location", "Run", "Savebutton", "Browse");
            //dataGridClass.addButtonColumn(dataGridView1, "Browse load location", "Run", "Loadbutton", "Browse");
            //dataGridClass.addButtonColumn(dataGridView1, "Run Report", "Run", "ReportButton", "Run");
           // dataGridClass.SetColumnsOrder(dataGridView1, "Stored_Procedure", "Save_location", "Browse save location", "Load_location", "Browse load location");
            dataGridClass.DisableTableSorting(dataGridView1);

        }
 
        //Save Data button
        private void button2_Click(object sender, EventArgs e)
        {
            //SAVES THE DATASET.

            //what is this (dataset)
            DataSet ds = (DataSet)dataGridView1.DataSource;
            ds.WriteXml(XMLData.getFilePath());
        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            //writes the string to the console.
           
            // this enables you to access custom generated buttons,  DataGridViewCellEventArgs is passing information about the location of the pressed cell.
            // e.columnindex and e.rowindex can be used to access the specific cell.
            try
            {
                richTextBox1.AppendText("\r\n" + dataGridView1[e.ColumnIndex, e.RowIndex].ToString()!);
                //first coolumn is browse for save location column
                if (e.ColumnIndex == 2)
                    {
                        SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
                        DialogResult result = SaveFileDialog1.ShowDialog(); // Show the dialog.

                        if (result == DialogResult.OK) // Test result.
                        {

                            string file = SaveFileDialog1.FileName;
                            dataGridView1[1, e.RowIndex].Value = file;

                        }
                    }
                else if (e.ColumnIndex == 4)
                {
                    OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
                    DialogResult result = OpenFileDialog1.ShowDialog(); // Show the dialog.

                    if (result == DialogResult.OK) // Test result.
                    {
                        string file = OpenFileDialog1.FileName;
                        dataGridView1[3, e.RowIndex].Value = file;
                    }
                }
                else if (e.ColumnIndex == 5)
                    {
                        var returnedDT = SQLAcess.SQLtoDataTable(dataGridView1[2, e.RowIndex].Value.ToString()!);

                        var loadstring = @"C:\Users\harrisonde\OneDrive - GXO\Desktop\Test\Test.xlsx";
                        var savestring = dataGridView1[3, e.RowIndex].Value.ToString()!;
                        var workbookstring = "test";

                        GXOMIClassLibrary.My_DataTable_Extensions.ExportToExcelDetailed(returnedDT, loadstring, workbookstring, savestring);
                    }


            }
            catch (ArgumentOutOfRangeException)
            {
                richTextBox1.AppendText("Argument out of range");
            }

        }

        private void QuerySave_Load(object sender, EventArgs e)
        {

            //load buttons.
            DataGridClass dataGridClass = new DataGridClass();
            dataGridClass.LoadButtons(dataGridView1);
        }
    }
}