using System.Data;
using System.Data.OleDb;



namespace excellreader
{
    internal class excellrdr
    {
        private DataTable ConvertExcelltoDataTable(string path)
        {
            string conStr = "";
            string FilePath = path;
            string Extension = ".xlsx";
            switch (Extension)
            {
                case ".xls":
                    conStr = $"Provider = Microsoft.ACE.OLEDB.12.0; Data Source ={FilePath}; Excel 12.0 Xml; HDR = YES";
                    break;
                case ".xlsx":
                    conStr = $"Provider = Microsoft.ACE.OLEDB.12.0; Data Source ={FilePath}; Excel 12.0 Xml; HDR = YES";
                    break;

            }
            conStr = String.Format(conStr, FilePath, "Yes");
            OleDbConnection connExcel = new OleDbConnection(conStr);
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter oda = new OleDbDataAdapter();

            DataTable dt = new DataTable();
            cmdExcel.Connection = connExcel;
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
        public DataSet GetAllfiles(string filepath)
        {
            DirectoryInfo d = new DirectoryInfo(filepath);
            FileInfo[] Files = d.GetFiles("*.xlsx");

            DataSet ds = new DataSet();
            
            foreach (FileInfo file in Files)
            {
                string path = filepath+"/"+file.Name;
                ds.Tables.Add(ConvertExcelltoDataTable(path));
            }
            return ds;
        }
    }

}
