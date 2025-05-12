using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using StrykerDemo;

namespace StrykerDemo
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            Load += Form2_Load;
        }

        private ComboBox cbProduct;
        private ComboBox cbCustomer;
        private Button btnReset;
        private DataGridView dgv;
        private DataTable fullData;

        private void Form2_Load(object sender, EventArgs e)
        {
            string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "Case_Study_Data.xlsx");
            if (!File.Exists(excelPath))
            {
                MessageBox.Show("Excel file not found: " + excelPath);
                return;
            }

            fullData = DataLoader.LoadMergedSalesWithProducts(excelPath);
            if (fullData == null)
            {
                MessageBox.Show("Failed to load data.");
                return;
            }

            cbProduct = new ComboBox
            {
                Left = 10,
                Top = 10,
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            cbProduct.Items.Add("All Products");
            cbProduct.Items.AddRange(fullData.AsEnumerable()
                .Select(r => r.Field<string>("ProductName"))
                .Where(p => !string.IsNullOrEmpty(p))
                .Distinct()
                .OrderBy(p => p)
                .Cast<object>()
                .ToArray());
            cbProduct.SelectedIndex = 0;
            Controls.Add(cbProduct);

            cbCustomer = new ComboBox
            {
                Left = 220,
                Top = 10,
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            cbCustomer.Items.Add("All Product Numbers");
            cbCustomer.Items.AddRange(fullData.AsEnumerable()
                .Select(r => r.Field<string>("ProductNumber"))
                .Where(c => !string.IsNullOrEmpty(c))
                .Distinct()
                .OrderBy(c => c)
                .Cast<object>()
                .ToArray());
            cbCustomer.SelectedIndex = 0;
            Controls.Add(cbCustomer);

            // Reset button
            btnReset = new Button
            {
                Text = "Reset Filters",
                Left = 440,
                Top = 10,
                Width = 120
            };
            btnReset.Click += (s, ev) => ResetFilters();
            Controls.Add(btnReset);

            // Table
            dgv = new DataGridView
            {
                Name = "dataGridView1",
                Top = 50,
                Left = 10,
                Width = 760,
                Height = 380,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                DataSource = fullData
            };
            Controls.Add(dgv);

            cbProduct.SelectedIndexChanged += FilterData;
            cbCustomer.SelectedIndexChanged += FilterData;

            SalesModule sales = new SalesModule();
            sales.Initialize(this, fullData);
        }

        private void FilterData(object sender, EventArgs e)
        {
            if (dgv == null || fullData == null)
                return;

            string selectedProduct = cbProduct?.SelectedItem?.ToString() ?? "All Products";
            string selectedCustomer = cbCustomer?.SelectedItem?.ToString() ?? "All Product Numbers";

            var filtered = fullData.AsEnumerable().Where(row =>
                (selectedProduct == "All Products" || row.Field<string>("ProductName") == selectedProduct) &&
                (selectedCustomer == "All Product Numbers" || row.Field<string>("ProductNumber") == selectedCustomer)
            );

            dgv.DataSource = filtered.Any() ? filtered.CopyToDataTable() : fullData.Clone();
        }

        private void ResetFilters()
        {
            cbProduct.SelectedIndex = 0;
            cbCustomer.SelectedIndex = 0;
            dgv.DataSource = fullData;
        }
    }
}
