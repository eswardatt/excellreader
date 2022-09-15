using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;


namespace excellreader
{
    internal class excellrdr
    {
        public DataTable GetSinglefile()
        {
            string conStr = "";
            string FilePath = "D:/Core";
            string Extension = ".xlsx";
            switch (Extension)
            {
                case ".xls":
                    conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Mode=ReadWrite;Extended Properties=Excel 12.0 Xml;HDR={1};IMEX=1" ;//ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                    break;
                case ".xlsx":
                    // conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Mode=ReadWrite;Extended Properties=Excel 12.0 Xml;HDR={1};IMEX=1";//ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                    conStr = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source =D:/Core/book.xlsx; Excel 12.0 Xml; HDR = YES";
                    break;

            }
            conStr = String.Format(conStr, FilePath, "Yes");
            OleDbConnection connExcel = new OleDbConnection(conStr);
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter oda = new OleDbDataAdapter();

            DataTable dt = new DataTable();

            cmdExcel.Connection = connExcel;
            //Get the name of First Sheet

            connExcel.Open();

            DataTable dtExcelSchema;

            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();



            connExcel.Close();
            //Read Data from First Sheet

            connExcel.Open();

            cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";


            //oda.TableMappings.Add("Table", "TestTable");

            oda.SelectCommand = cmdExcel;

            oda.Fill(dt);

            connExcel.Close();
            return dt;


        }

        DataTable GetGroupdata(string filename)
        {
            string conStr = "";
            string FilePath = "D:/Core/"+ filename;
            string Extension = ".xlsx";
            switch (Extension)
            {
                case ".xls":
                    conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Mode=ReadWrite;Extended Properties=Excel 12.0 Xml;HDR={1};IMEX=1";//ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                    break;
                case ".xlsx":
                    // conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Mode=ReadWrite;Extended Properties=Excel 12.0 Xml;HDR={1};IMEX=1";//ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                    conStr = $"Provider = Microsoft.ACE.OLEDB.12.0; Data Source ={FilePath}; Excel 12.0 Xml; HDR = YES";
                    break;

            }
            conStr = String.Format(conStr, FilePath, "Yes");
            OleDbConnection connExcel = new OleDbConnection(conStr);
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter oda = new OleDbDataAdapter();

            DataTable dt = new DataTable();

            cmdExcel.Connection = connExcel;
            //Get the name of First Sheet

            connExcel.Open();

            DataTable dtExcelSchema;

            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();



            connExcel.Close();
            //Read Data from First Sheet

            connExcel.Open();

            cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";


            //oda.TableMappings.Add("Table", "TestTable");

            oda.SelectCommand = cmdExcel;

            oda.Fill(dt);

            connExcel.Close();
            return dt;


        }
        public DataSet GetAllfiles()
        {
            DirectoryInfo d = new DirectoryInfo(@"D:\Core"); //Assuming Test is your Folder
            FileInfo[] Files = d.GetFiles("*.xlsx"); //Getting Text files
            
            DataSet ds = new DataSet();
            foreach (FileInfo file in Files)
            {
                ds.Tables.Add(GetGroupdata(file.Name));
            }
           return ds;
        }
    }
 
}
