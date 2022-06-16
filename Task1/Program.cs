using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using ClosedXML.Excel;
using System.Xml;
using System.IO;

namespace Task1
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = args[0];
            DataTable dt = new DataTable();
            var wbook = new XLWorkbook(fileName);
            var ws1 = wbook.Worksheet(1);
            int ColumnCount = ws1.ColumnsUsed().Count();
            int RowsCount = ws1.RowsUsed().Count();
            //Console.WriteLine(ColumnCount+" "+RowsCount);


            //Reading First Row as Column Hearders
            for(int i=1;i<ColumnCount;i++)
            {
                string value = ws1.Row(1).Cell(i).Value.ToString();
                //Console.WriteLine(value);
                dt.Columns.Add(value, typeof(string));
            }

            //Reading Second Row as Row
            DataRow dr = dt.NewRow();
            for (int i = 1; i < ColumnCount; i++)
            {
                string value = ws1.Row(2).Cell(i).Value.ToString();
                //Console.WriteLine(value);
                dr[i - 1] = value;
            }
            dt.Rows.Add(dr);


            DataRow dr1 = dt.NewRow();
            for (int i = 1; i < ColumnCount; i++)
            {
                string value = ws1.Row(3).Cell(i).Value.ToString();
                //Console.WriteLine(value);
                dr1[i - 1] = value;
            }
            dt.Rows.Add(dr1);


            //foreach (DataColumn column in dt.Columns)
            //{
            //    //Console.Write( column.ColumnName);

            //}

            //foreach (DataRow row in dt.Rows)
            //{
            //    foreach (DataColumn column in dt.Columns)
            //    {
            //        Console.WriteLine(row[column]);
            //    }
            //}


            //Printing data of datatable
            //for (int j = 0; j < dt.Rows.Count; j++)
            //{
            //    for (int i = 0; i < dt.Columns.Count; i++)
            //    {
            //        Console.Write(dt.Columns[i].ColumnName + " ");
            //        Console.WriteLine(dt.Rows[j].ItemArray[i] + " ");

            //    }
            //}


            string JSONresult;
            JSONresult = JsonConvert.SerializeObject(dt);
            Console.Write(JSONresult);
            string time=DateTime.Now.ToString("yyyyMMddHHmmssffff");
            Console.WriteLine(time);
            //System.IO.File.WriteAllText(@"C:\Inputs\"+time+"path.txt", JSONresult);

            Console.WriteLine("\n\n\nXML placeholder replacement started....");

            //XmlDocument itemDoc = new XmlDocument();
            //itemDoc.Load(@"C:\Inputs\DemoTemplate.xml");
            //Console.WriteLine(itemDoc.DocumentElement.ChildNodes.Count);
            string text = File.ReadAllText(@"C:\Inputs\DemoTemplate.xml");
            
            text=text.Replace("##SiteId##","demo");
            //Console.WriteLine(text);
            System.IO.File.WriteAllText(@"C:\Inputs\" + time + "UpdatedXML.xml", text);

            Console.ReadKey();

        }

    }
}
