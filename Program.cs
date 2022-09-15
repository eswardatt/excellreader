// See https://aka.ms/new-console-template for more information
using excellreader;
using System.Data;

Console.WriteLine("Hello, World!");
excellrdr excellreader = new excellrdr();

DataTable dt = excellreader.Get();
int i = dt.Rows.Count;
Console.WriteLine(i);
foreach (DataRow row in dt.Rows)
{
  Console.WriteLine("Id : {0} && Name : {1}",  row["id"].ToString(), row["name"].ToString());
}