
using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace StrykerDemo
{
    public class SalesModule
    {
        private DataGridView dgvSales;
        private Chart chartSales;
        private ComboBox cbProduct;
        private DateTimePicker dtStart, dtEnd;
        private Button btnFilter, btnReset;
        private DataTable salesData;

        public void Initialize(Control parent, DataTable salesData)
        {
            this.salesData = salesData;

            dtStart = new DateTimePicker { Left = 10, Top = 10, Width = 120 };
            dtEnd = new DateTimePicker { Left = 140, Top = 10, Width = 120 };
            cbProduct = new ComboBox { Left = 270, Top = 10, Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };
            btnFilter = new Button { Text = "Filter", Left = 480, Top = 10 };
            btnReset = new Button { Text = "Reset", Left = 560, Top = 10 };

            dgvSales = new DataGridView
            {
                Left = 10,
                Top = 50,
                Width = 760,
                Height = 300,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            chartSales = new Chart
            {
                Left = 10,
                Top = 360,
                Width = 760,
                Height = 300
            };
            ChartArea chartArea = new ChartArea("SalesArea");
            chartSales.ChartAreas.Add(chartArea);
            chartSales.Series.Add(new Series("Sales")
            {
                ChartType = SeriesChartType.Line,
                XValueType = ChartValueType.Date
            });

            parent.Controls.AddRange(new Control[] { dtStart, dtEnd, cbProduct, btnFilter, btnReset, dgvSales, chartSales });

            // Populate product combo
            cbProduct.Items.Add("All Products");
            cbProduct.Items.AddRange(salesData.AsEnumerable()
                .Select(r => r.Field<string>("SalesOrderNumber"))
                .Where(p => !string.IsNullOrEmpty(p))
                .Distinct()
                .OrderBy(p => p)
                .Cast<object>()
                .ToArray());
            cbProduct.SelectedIndex = 0;

            // Set default date range
            dtStart.Value = salesData.AsEnumerable().Min(r => r.Field<DateTime>("OrderDate"));
            dtEnd.Value = salesData.AsEnumerable().Max(r => r.Field<DateTime>("OrderDate"));

            btnFilter.Click += FilterAndDisplay;
            btnReset.Click += (s, e) =>
            {
                cbProduct.SelectedIndex = 0;
                dtStart.Value = salesData.AsEnumerable().Min(r => r.Field<DateTime>("OrderDate"));
                dtEnd.Value = salesData.AsEnumerable().Max(r => r.Field<DateTime>("OrderDate"));
                FilterAndDisplay(null, null);
            };

            FilterAndDisplay(null, null);
        }

        private void FilterAndDisplay(object sender, EventArgs e)
        {
            DateTime start = dtStart.Value.Date;
            DateTime end = dtEnd.Value.Date;
            string selectedProduct = cbProduct.SelectedItem.ToString();

            var filteredRows = salesData.AsEnumerable().Where(row =>
                row.Field<DateTime>("OrderDate") >= start &&
                row.Field<DateTime>("OrderDate") <= end &&
                (selectedProduct == "All Products" || row.Field<string>("ProductName") == selectedProduct)
            );

            DataTable filtered = filteredRows.Any() ? filteredRows.CopyToDataTable() : salesData.Clone();
            dgvSales.DataSource = filtered;

            // Chart
            var salesByDate = filtered.AsEnumerable()
                .GroupBy(r => r.Field<DateTime>("OrderDate"))
                .Select(g => new
                {
                    Date = g.Key,
                    Total = g.Sum(r => r.Field<double>("LineTotal"))
                })
                .OrderBy(x => x.Date)
                .ToList();

            chartSales.Series["Sales"].Points.Clear();
            foreach (var item in salesByDate)
            {
                chartSales.Series["Sales"].Points.AddXY(item.Date, item.Total);
            }
        }
    }
}
