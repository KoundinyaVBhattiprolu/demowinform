using System;
using System.Data;
using System.IO;
using System.Linq;
using ExcelDataReader;

namespace StrykerDemo
{
    public static class DataLoader
    {
        public static DataTable LoadMergedSalesWithProducts(string excelPath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            var result = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true }
            });

            var header = result.Tables["SalesOrderHeader"];
            var detail = result.Tables["SalesOrderDetail"];
            var product = result.Tables["Product"];

            // Rename columns
            if (product.Columns.Contains("Name"))
                product.Columns["Name"].ColumnName = "ProductName";

            // Merge: SalesOrderDetail + SalesOrderHeader + Product
            var merged = from d in detail.AsEnumerable()
                         join h in header.AsEnumerable()
                           on d.Field<double>("SalesOrderID") equals h.Field<double>("SalesOrderID")
                         join p in product.AsEnumerable()
                           on d.Field<double>("ProductID") equals p.Field<double>("ProductID")
                         select new
                         {
                             SalesOrderID = d.Field<double>("SalesOrderID"),
                             SalesOrderDetailID = d.Field<double>("SalesOrderDetailID"),
                             SalesOrderNumber = h.Field<string>("SalesOrderNumber"),
                             ProductID = d.Field<double>("ProductID"),
                             ProductName = p.Field<string>("ProductName"),
                             ProductNumber = p.Field<string>("ProductNumber"),
                             OrderQty = d.Field<double>("OrderQty"),
                             LineTotal = d.Field<double>("LineTotal"),
                             OrderDate = h.Field<DateTime>("OrderDate")
                         };

            // Build the final DataTable
            var final = new DataTable();
            final.Columns.Add("SalesOrderID", typeof(double));
            final.Columns.Add("SalesOrderDetailID", typeof(double));
            final.Columns.Add("SalesOrderNumber", typeof(string));
            final.Columns.Add("ProductID", typeof(double));
            final.Columns.Add("ProductName", typeof(string));
            final.Columns.Add("ProductNumber", typeof(string));
            final.Columns.Add("OrderQty", typeof(double));
            final.Columns.Add("LineTotal", typeof(double));
            final.Columns.Add("OrderDate", typeof(DateTime));

            foreach (var row in merged)
            {
                final.Rows.Add(row.SalesOrderID, row.SalesOrderDetailID, row.SalesOrderNumber, row.ProductID,
                               row.ProductName, row.ProductNumber,
                               row.OrderQty, row.LineTotal, row.OrderDate);
            }

            return final;
        }
    }
}
