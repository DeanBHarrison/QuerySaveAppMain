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
        // column indexs \\
        //0 - stored_procedure
        //1 - save location
        //2 - browse save
        //3 - load location
        //4 - browse load
        //5 - run
        //6 - settings

        public QuerySave()
        {
            InitializeComponent();
            //set Datasource and Datamember, both of these are required to return data.
            dataGridView1.DataSource = XMLData.ReturnXMLDataset(1);
            dataGridView1.DataMember = "Authors";

            ////add a button column and then add buttons to it
            DataGridClass dataGridClass = new DataGridClass();
            //dataGridClass.addButtonColumn(dataGridView1, "Browse save location", "Run", "Savebutton", "Browse");
            //dataGridClass.addButtonColumn(dataGridView1, "Browse load location", "Run", "Loadbutton", "Browse");
            //dataGridClass.addButtonColumn(dataGridView1, "Run Report", "Run", "ReportButton", "Run");
            //dataGridClass.SetColumnsOrder(dataGridView1, "Stored_Procedure", "Save_location", "Browse save location", "Load_location", "Browse load location");
            dataGridClass.DisableTableSorting(dataGridView1);

        }
 
        //Save Data button
        private void button2_Click(object sender, EventArgs e)
        {
            //SAVES THE DATASET.

            //what is this (dataset)
            DataSet ds = (DataSet)dataGridView1.DataSource;
            ds.WriteXml(XMLData.getFilePath(1));
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
                //chose load loaction
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
                //run SQL - save to location
                else if (e.ColumnIndex == 5)
                    {
                    // return SQL into datatable

                        var returnedDT = SQLAcess.SQLtoDataTable(dataGridView1[0, e.RowIndex].Value.ToString()!);

                    var SettingsDataset = XMLData.ReturnXMLDataset(2);
                    

                    richTextBox1.AppendText(SettingsDataset.Tables[0].Rows[0][0].ToString()!);

                    //open this item
                        var loadstring = dataGridView1[3, e.RowIndex].Value.ToString()!;
                    //find the string for the workbook to paste into
                    var workbookstring = SettingsDataset.Tables[0].Rows[e.RowIndex][0].ToString()!;              //get the fuckin string from the second xml dataset DO THIS
                    //save it to this location
                    var savestring = dataGridView1[1, e.RowIndex].Value.ToString()!;
                   

                        GXOMIClassLibrary.My_DataTable_Extensions.ExportToExcelDetailed(returnedDT, loadstring, workbookstring, savestring);
                    }
                else if (e.ColumnIndex == 6)
                {
                    Settings settingsform = new Settings();

                    settingsform.Show();
                }


            }
            catch (Exception ex)
            {
                richTextBox1.AppendText(ex.Message);
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