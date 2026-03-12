using OfficeOpenXml;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Windows;

namespace WpfApp1
{
    public partial class _4338_Galiev : Window
    {
        private string connectionString = @"Server=localhost\SQLEXPRESS;Database=Lab3DB;Trusted_Connection=True;";

        public _4338_Galiev()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // важно для EPPlus
        }

        private void InsertOrder(string orderCode, string clientCode, string services, string orderDate)
        {
            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string sql = @"INSERT INTO Orders (OrderCode, ClientCode, Services, OrderDate) 
                               VALUES (@code, @client, @services, @date)";
                using (var cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@code", orderCode);
                    cmd.Parameters.AddWithValue("@client", clientCode);
                    cmd.Parameters.AddWithValue("@services", services);
                    cmd.Parameters.AddWithValue("@date", orderDate);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private List<Order> GetAllOrders()
        {
            var list = new List<Order>();
            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                string sql = "SELECT * FROM Orders";
                using (var cmd = new SqlCommand(sql, conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        list.Add(new Order
                        {
                            Id = reader.GetInt32(0),
                            OrderCode = reader.GetString(1),
                            ClientCode = reader.GetString(2),
                            Services = reader.GetString(3),
                            OrderDate = reader.GetString(4)
                        });
                    }
                }
            }
            return list;
        }

        private void ImportData_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "Excel files|*.xlsx";
            dialog.Title = "Выберите файл 2.xlsx";

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    using (var package = new ExcelPackage(new FileInfo(dialog.FileName)))
                    {
                        var worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension.Rows;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            var orderCode = worksheet.Cells[row, 1].Text;
                            var clientCode = worksheet.Cells[row, 2].Text;
                            var services = worksheet.Cells[row, 3].Text;
                            var orderDate = worksheet.Cells[row, 4].Text;

                            InsertOrder(orderCode, clientCode, services, orderDate);
                        }
                    }

                    MessageBox.Show("Импорт данных завершён!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при импорте: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ExportData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var orders = GetAllOrders();
                var grouped = orders.GroupBy(o => o.OrderDate.Substring(0, 7));

                using (var package = new ExcelPackage())
                {
                    foreach (var group in grouped)
                    {
                        string sheetName = group.Key;
                        var worksheet = package.Workbook.Worksheets.Add(sheetName);

                        worksheet.Cells[1, 1].Value = "Id";
                        worksheet.Cells[1, 2].Value = "Код заказа";
                        worksheet.Cells[1, 3].Value = "Код клиента";
                        worksheet.Cells[1, 4].Value = "Услуги";

                        int row = 2;
                        foreach (var order in group)
                        {
                            worksheet.Cells[row, 1].Value = order.Id;
                            worksheet.Cells[row, 2].Value = order.OrderCode;
                            worksheet.Cells[row, 3].Value = order.ClientCode;
                            worksheet.Cells[row, 4].Value = order.Services;
                            row++;
                        }

                        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    }

                    var saveDialog = new SaveFileDialog();
                    saveDialog.Filter = "Excel files|*.xlsx";
                    saveDialog.FileName = "export.xlsx";
                    if (saveDialog.ShowDialog() == true)
                    {
                        package.SaveAs(new FileInfo(saveDialog.FileName));
                        MessageBox.Show("Экспорт завершён!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при экспорте: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    public class Order
    {
        public int Id { get; set; }
        public string OrderCode { get; set; }
        public string ClientCode { get; set; }
        public string Services { get; set; }
        public string OrderDate { get; set; }
    }
}