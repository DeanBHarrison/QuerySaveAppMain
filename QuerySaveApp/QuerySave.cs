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
            DataGridClass.DisableTableSorting(dataGridView1);
        }

        //Save Data button
        private void button2_Click(object sender, EventArgs e)
        {
            DataSet ds = (DataSet)dataGridView1.DataSource;
            ds.WriteXml(XMLData.getFilePath(1));
        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                richTextBox1.AppendText("\r\n" + dataGridView1[e.ColumnIndex, e.RowIndex].ToString()!);
                //first coolumn is browse for save location column
                if (((DataGridView)sender).Columns[e.ColumnIndex].DataPropertyName == "Browse_Save")
                    {
                        SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
                        DialogResult result = SaveFileDialog1.ShowDialog(); // Show the dialog.

                        if (result == DialogResult.OK) // Test result.
                        {
                            string file = SaveFileDialog1.FileName;
                        var index = dataGridView1.Columns["Save_Location"].Index;
                        // need to make the value of the column with the index name "load_location" at the same row index.
                        dataGridView1[index, e.RowIndex].Value = file;
                        }
                    }
                //chose load loaction
                else if (((DataGridView)sender).Columns[e.ColumnIndex].DataPropertyName == "Browse_load")
                {
                    OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
                    DialogResult result = OpenFileDialog1.ShowDialog(); // Show the dialog.

                    if (result == DialogResult.OK) // Test result.
                    {
                        string file = OpenFileDialog1.FileName;
                        var index = dataGridView1.Columns["Load_Location"].Index;
                        dataGridView1[index, e.RowIndex].Value = file;
                    }
                }
                //run SQL - save to location, workbook
                else if (((DataGridView)sender).Columns[e.ColumnIndex].DataPropertyName == "Run")
                    {
                    // return SQL into datatable
                    var returnedDT = SQLAcess.SQLtoDataTable(dataGridView1[0, e.RowIndex].Value.ToString()!);

                    //find item to open
                    string loadstring = DataGridClass.CellColumn(dataGridView1, "Load_Location", e.RowIndex);

                    //finds workbook to paste into. THIS METHOD looks to the second XML dataset on the other form.
                    //this can be used to return a cell straight from an XML dataset, skipping the requirement of it to be a datagridview first.
                    var SettingsDataset = XMLData.ReturnXMLDataset(2);       
                    var workbookstring = XMLData.returnXMLcellwithcolumnname(SettingsDataset, "Data_Dump_Worksheet_name", e.RowIndex);

                    //find location to save it
                    string savestring = DataGridClass.CellColumn(dataGridView1, "Save_location", e.RowIndex);

                    //execute export to excel, with the locations saved from above.
                    GXOMIClassLibrary.My_DataTable_Extensions.ExportToExcelDetailed(returnedDT, loadstring, workbookstring, savestring);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    }
                else if (((DataGridView)sender).Columns[e.ColumnIndex].DataPropertyName == "Settings")
                {
                    Settings settingsform = new Settings();
                    settingsform.Show();
                }
        }
            catch (Exception ex)
            {
               richTextBox1.AppendText(ex.Message);
                //Your Excel object).Application.Interactive = false;  try this
            }
}

        private void QuerySave_Load(object sender, EventArgs e)
        {
            //load buttons.
           // DataGridClass dataGridClass = new DataGridClass();
            DataGridClass.LoadButtons(dataGridView1);
        }
    }
}