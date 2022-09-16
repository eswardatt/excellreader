// See https://aka.ms/new-console-template for more information
using excellreader;
using System.Data;

excellrdr excellreader = new excellrdr();
Console.WriteLine("Enter path:");
string path = Console.ReadLine().ToString();
DataSet ds = excellreader.ReadAllExcellFiles(path);
//DataSet ds = excellreader.ReadAllExcellFiles(@"D:\Core");
int tablescount = ds.Tables.Count;
if (tablescount > 0)
{
    for (int i = 0; i < tablescount; i++)
    {
        DataTable dt = ds.Tables[i];
        Console.WriteLine(ds.Tables[i].TableName);
        foreach (DataRow row in dt.Rows)
            Console.WriteLine("Id : {0} && Name : {1}", row["id"].ToString(), row["name"].ToString());
    }
}
else
    Console.WriteLine("No Tables found");