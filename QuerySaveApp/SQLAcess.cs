using Excel = Microsoft.Office.Interop.Excel;
// excel object library is required to use the above, not the fucking shitty interlop piece of shit.
using System.Reflection;//fuck knows what this one is being used for.
// gotta use system not microsoft or it errors on using a certificate for some reason.
using System.Data.SqlClient;
using System;
using System.Data;

namespace QuerySaveApp
{
    internal class SQLAcess
    {
        public static DataTable SQLtoDataTable(string StoredProcedureName)
        {
            //set connection
            SqlConnection connection = new SqlConnection("Data Source=10.8.232.50; Initial Catalog=WMSMIS;User ID=gxo.reports.orphan;Password=yAQexG*@myj8$u");
            //set query
            string queryString = "exec " + StoredProcedureName;
            //create a command object
            SqlCommand command = new SqlCommand(queryString, connection);
            //open connection to SQL server
            connection.Open();

            //Create data adaptor
            SqlDataAdapter dataAdapter = new SqlDataAdapter(command);

            //create a data table
            DataTable dataTable = new DataTable();

            dataAdapter.Fill(dataTable);
            return dataTable;
        }
    }
}
