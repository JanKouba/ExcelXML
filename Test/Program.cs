using System;
using System.Collections.Generic;
using System.Text;
using ExcelXML;
using System.Drawing;
using System.Diagnostics;
using System.Data;

namespace Test2
{
    class Program
    {
        static void Main(string[] args)
        {
            //database db = new database();


            //XML.Header.TextColor = XML.Header.TextColor;

            DataTable dt1 = new DataTable("Sheet1");
            DataTable dt2 = new DataTable("Sheet2");

            DataTable[] dtSource = { dt2 };

            dt1.Columns.Add("c1");
            dt1.Columns.Add("c2");
            dt1.Columns.Add("c3");
            
            dt2.Columns.Add("c1");
            dt2.Columns.Add("c02");
            dt2.Columns.Add("c03");

            DataRow dr = dt1.NewRow();
            
            dr[0] = "0";
            dr[1] = "1";
            dr[2] = "2";
            dt1.Rows.Add(dr);
            
            dr = dt2.NewRow();

            dr[0] = "3";
            dr[1] = "8";
            dr[2] = "5";

            dt2.Rows.Add(dr);

            DataView dv = Common.ColumnFilter(dt1, new string[] { "c2","c1"});

            dt1 = null;
        }
    }
}
