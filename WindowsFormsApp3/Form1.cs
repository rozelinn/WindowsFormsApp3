using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace WarehousePathOptimization
{
    public partial class Form1 : Form
    {
        private List<Order> orders;
        private Dictionary<string, Point> locationMap;
        private Point depotEntrance; // Depo giriş noktası

        public Form1(ExcelPackage excelPackage)
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeLocationMap();
            InitializeDepotEntrance(); // Depo giriş noktasını başlat
            InitializeOrderData();
            PopulateProductList();
        }

        private void InitializeDepotEntrance()
        {
            // Depo giriş noktası için örnek koordinatlar
            depotEntrance = new Point(65, 620); // Giriş noktası (örnek)
        }

        private void InitializeLocationMap()
        {
            locationMap = new Dictionary<string, Point>
            {
                { "5H/1", new Point(505, 280) },
                { "5H/2", new Point(505, 295) },
                { "5H/3", new Point(505, 310) },
                { "5H/4", new Point(505, 325) },
                { "5H/5", new Point(505, 345) },
                { "5H/6", new Point(505, 365) },
                { "5H/7", new Point(505, 385) },
                { "5H/8", new Point(505, 405) },
                { "5H/9", new Point(505, 420) },
                { "5H/10", new Point(505, 440) },
                { "5H/11", new Point(505, 455) },
                { "5H/12", new Point(505, 470) },

                { "6H/1", new Point(450, 280) },
                { "6H/2", new Point(450, 295) },
                { "6H/3", new Point(450, 310) },
                { "6H/4", new Point(450, 325) },
                { "6H/5", new Point(450, 345) },
                { "6H/6", new Point(450, 365) },
                { "6H/7", new Point(450, 385) },
                { "6H/8", new Point(450, 405) },
                { "6H/9", new Point(450, 420) },
                { "6H/10", new Point(450, 440) },
                { "6H/11", new Point(450, 455) },
                { "6H/12", new Point(450, 470) },

                { "7H/1", new Point(400, 10) },
                { "7H/2", new Point(400, 30) },
                { "7H/3", new Point(400, 55) },
                { "7H/4", new Point(400, 70) },
                { "7H/5", new Point(400, 90) },
                { "7H/6", new Point(400, 105) },
                { "7H/7", new Point(400, 120) },
                { "7H/8", new Point(400, 135) },
                { "7H/9", new Point(400, 150) },
                { "7H/10", new Point(400, 165) },
                { "7H/11", new Point(400, 180) },
                { "7H/12", new Point(400, 280) },
                { "7H/13", new Point(400, 295) },
                { "7H/14", new Point(400, 310) },
                { "7H/15", new Point(400, 325) },
                { "7H/16", new Point(400, 345) },
                { "7H/17", new Point(400, 365) },
                { "7H/18", new Point(400, 385) },
                { "7H/19", new Point(400, 405) },
                { "7H/20", new Point(400, 420) },
                { "7H/21", new Point(400, 440) },
                { "7H/22", new Point(400, 455) },
                { "7H/23", new Point(400, 470) },

                { "8H/1", new Point(350, 10) },
                { "8H/2", new Point(350, 30) },
                { "8H/3", new Point(350, 55) },
                { "8H/4", new Point(350, 70) },
                { "8H/5", new Point(350, 90) },
                { "8H/6", new Point(350, 105) },
                { "8H/7", new Point(350, 120) },
                { "8H/8", new Point(350, 135) },
                { "8H/9", new Point(350, 150) },
                { "8H/10", new Point(350, 165) },
                { "8H/11", new Point(350, 180) },
                { "8H/12", new Point(350, 280) },
                { "8H/13", new Point(350, 295) },
                { "8H/14", new Point(350, 310) },
                { "8H/15", new Point(350, 325) },
                { "8H/16", new Point(350, 345) },
                { "8H/17", new Point(350, 365) },
                { "8H/18", new Point(350, 385) },
                { "8H/19", new Point(350, 405) },
                { "8H/20", new Point(350, 420) },
                { "8H/21", new Point(350, 440) },
                { "8H/22", new Point(350, 455) },
                { "8H/23", new Point(350, 470) },

                { "9H/1", new Point(300, 10) },
                { "9H/2", new Point(300, 30) },
                { "9H/3", new Point(300, 55) },
                { "9H/4", new Point(300, 70) },
                { "9H/5", new Point(300, 90) },
                { "9H/6", new Point(300, 105) },
                { "9H/7", new Point(300, 120) },
                { "9H/8", new Point(300, 135) },
                { "9H/9", new Point(300, 150) },
                { "9H/10", new Point(300, 165) },
                { "9H/11", new Point(300, 180) },
                { "9H/12", new Point(300, 215) },
                { "9H/13", new Point(300, 230) },
                { "9H/14", new Point(300, 250) },
                { "9H/15", new Point(300, 280) },
                { "9H/16", new Point(300, 295) },
                { "9H/17", new Point(300, 310) },
                { "9H/18", new Point(300, 325) },
                { "9H/19", new Point(300, 345) },
                { "9H/20", new Point(300, 365) },
                { "9H/21", new Point(300, 385) },
                { "9H/22", new Point(300, 405) },
                { "9H/23", new Point(300, 420) },
                { "9H/24", new Point(300, 440) },
                { "9H/25", new Point(300, 455) },
                { "9H/26", new Point(300, 470) },



                { "10H/1", new Point(250, 10) },
                { "10H/2", new Point(250, 30) },
                { "10H/3", new Point(250, 55) },
                { "10H/4", new Point(250, 70) },
                { "10H/5", new Point(250, 90) },
                { "10H/6", new Point(250, 105) },
                { "10H/7", new Point(250, 120) },
                { "10H/8", new Point(250, 135) },
                { "10H/9", new Point(250, 150) },
                { "10H/10", new Point(250, 165) },
                { "10H/11", new Point(250, 180) },
                { "10H/12", new Point(250, 215) },
                { "10H/13", new Point(250, 230) },
                { "10H/14", new Point(250, 250) },
                { "10H/15", new Point(250, 280) },
                { "10H/16", new Point(250, 295) },
                { "10H/17", new Point(250, 310) },
                { "10H/18", new Point(250, 325) },
                { "10H/19", new Point(250, 345) },
                { "10H/20", new Point(250, 365) },
                { "10H/21", new Point(250, 385) },
                { "10H/22", new Point(250, 405) },
                { "10H/23", new Point(250, 420) },
                { "10H/24", new Point(250, 440) },
                { "10H/25", new Point(250, 455) },
                { "10H/26", new Point(250, 470) },


                { "11H/1", new Point(200, 10) },
                { "11H/2", new Point(200, 30) },
                { "11H/3", new Point(200, 55) },
                { "11H/4", new Point(200, 70) },
                { "11H/5", new Point(200, 90) },
                { "11H/6", new Point(200, 105) },
                { "11H/7", new Point(200, 120) },
                { "11H/8", new Point(200, 135) },
                { "11H/9", new Point(200, 150) },
                { "11H/10", new Point(200, 165) },
                { "11H/11", new Point(200, 180) },
                { "11H/12", new Point(200, 215) },
                { "11H/13", new Point(200, 230) },
                { "11H/14", new Point(200, 250) },
                { "11H/15", new Point(200, 280) },
                { "11H/16", new Point(200, 295) },
                { "11H/17", new Point(200, 310) },
                { "11H/18", new Point(200, 325) },
                { "11H/19", new Point(200, 345) },
                { "11H/20", new Point(200, 365) },
                { "11H/21", new Point(200, 385) },
                { "11H/22", new Point(200, 405) },
                { "11H/23", new Point(200, 420) },
                { "11H/24", new Point(200, 440) },
                { "11H/25", new Point(200, 455) },
                { "11H/26", new Point(200, 470) },


                { "12H/1", new Point(120, 10) },
                { "12H/2", new Point(120, 30) },
                { "12H/3", new Point(120, 55) },
                { "12H/4", new Point(120, 65) },
                { "12H/5", new Point(120, 90) },
                { "12H/6", new Point(120, 105) },
                { "12H/7", new Point(120, 120) },
                { "12H/8", new Point(120, 135) },
                { "12H/9", new Point(120, 150) },
                { "12H/10", new Point(120, 165) },
                { "12H/11", new Point(120, 180) },
                { "12H/12", new Point(120, 215) },
                { "12H/13", new Point(120, 230) },
                { "12H/14", new Point(120, 250) },
                { "12H/15", new Point(120, 280) },
                { "12H/16", new Point(120, 295) },
                { "12H/17", new Point(120, 310) },
                { "12H/18", new Point(120, 325) },
                { "12H/19", new Point(120, 345) },
                { "12H/20", new Point(120, 365) },
                { "12H/21", new Point(120, 385) },
                { "12H/22", new Point(120, 405) },
                { "12H/23", new Point(120, 420) },
                { "12H/24", new Point(120, 440) },
                { "12H/25", new Point(120, 455) },
                { "12H/26", new Point(120, 470) },

                { "13H/1", new Point(450, 100) },
                { "13H/2", new Point(450, 150) },
                { "13H/3", new Point(450, 200) }
            };
        }

        private void InitializeOrderData()
        {
            orders = new List<Order>();
            string excelFilePath = @"C:\\Users\\hp\\source\\repos\\WindowsFormsApp3\\WindowsFormsApp3\\depo_tum_hucreler_urunler.xlsx";

            try
            {
                FileInfo fileInfo = new FileInfo(excelFilePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        var productName = worksheet.Cells[row, 1].Text;
                        var productIdStr = productName.Split(' ')[1];
                        int productId = int.TryParse(productIdStr, out var pid) ? pid : 0;

                        int cell = int.TryParse(worksheet.Cells[row, 2].Text, out var c) ? c : 0;
                        var corridorText = worksheet.Cells[row, 3].Text;
                        int rack = int.TryParse(worksheet.Cells[row, 4].Text, out var r) ? r : 0;

                        if (productId > 0 && cell > 0 && rack > 0 && !string.IsNullOrEmpty(corridorText))
                        {
                            orders.Add(new Order(productId, corridorText, rack, cell));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel dosyası okunamadı: " + ex.Message);
            }
        }

        private void PopulateProductList()
        {
            var productIds = orders.Select(o => o.ProductId).Distinct().ToList();
            foreach (var id in productIds)
            {
                checkedListBoxProducts.Items.Add($"Ürün {id}");
            }
        }

        private void buttonCalculate_Click(object sender, EventArgs e)
        {
            var selectedProducts = checkedListBoxProducts.CheckedItems
                .Cast<string>()
                .Select(item => int.Parse(item.Split(' ')[1]))
                .ToList();

            if (!selectedProducts.Any())
            {
                MessageBox.Show("Lütfen en az bir ürün seçin!");
                return;
            }

            var optimizedPath = CalculateOptimizedPath(selectedProducts, out double totalDistance);

            textBoxOutput.Clear();
            textBoxOutput.AppendText("En Yakın Konumlara Göre Sipariş Toplama Sırası:\n");
            textBoxOutput.AppendText("---------------------------------------------\n");
            int step = 1;
            foreach (var order in optimizedPath)
            {
                textBoxOutput.AppendText(
                    $"Adım {step++}: Ürün {order.ProductId} --> Koridor {order.Corridor}, Raf {order.Rack}, Hücre {order.Cell}\n");
            }
            textBoxOutput.AppendText("---------------------------------------------\n");
            textBoxOutput.AppendText($"Toplam Kat Edilen Mesafe: {totalDistance:F2} metre\n");

            DrawOrderLocations(optimizedPath);
        }

        private List<Order> CalculateOptimizedPath(List<int> selectedProducts, out double totalDistance)
        {
            var remainingOrders = orders
                .Where(o => selectedProducts.Contains(o.ProductId))
                .GroupBy(o => o.ProductId)
                .Select(g => g.OrderBy(o => ExtractNumberFromString(o.Corridor)).First())
                .ToList();

            var path = new List<Order>();
            totalDistance = 0;
            Point currentLocation = depotEntrance; // İlk konum olarak depo girişi kullanılıyor

            while (remainingOrders.Any())
            {
                Order nextOrder = null;
                double minDistance = double.MaxValue;

                foreach (var order in remainingOrders)
                {
                    Point orderLocation = locationMap[$"{order.Corridor}/{order.Rack}"];
                    double distance = CalculateDistance(currentLocation, orderLocation);

                    if (distance < minDistance)
                    {
                        minDistance = distance;
                        nextOrder = order;
                    }
                }

                if (nextOrder != null)
                {
                    path.Add(nextOrder);
                    remainingOrders.Remove(nextOrder);
                    totalDistance += minDistance;
                    currentLocation = locationMap[$"{nextOrder.Corridor}/{nextOrder.Rack}"];
                }
            }

            return path;
        }

        private double CalculateDistance(Point p1, Point p2)
        {
            return Math.Sqrt(
                Math.Pow(p1.X - p2.X, 2) +
                Math.Pow(p1.Y - p2.Y, 2)
            );
        }

        private int ExtractNumberFromString(string input)
        {
            string numberPart = new string(input.Where(char.IsDigit).ToArray());
            return int.TryParse(numberPart, out int result) ? result : 0;
        }

        private void DrawOrderLocations(List<Order> optimizedPath)
        {
            var map = new Bitmap(pictureBoxMap.Image);
            using (Graphics g = Graphics.FromImage(map))
            {
                foreach (var order in optimizedPath)
                {
                    string locationKey = $"{order.Corridor}/{order.Rack}";
                    if (locationMap.TryGetValue(locationKey, out Point coordinates))
                    {
                        int x = coordinates.X;
                        int y = coordinates.Y;

                        g.FillEllipse(Brushes.Red, x, y, 15, 15);
                        g.DrawString($"Ürün {order.ProductId}", new Font("Arial", 8), Brushes.Black, x, y - 10);
                    }
                }

                // Depo giriş noktasını mavi bir daire ile işaretleyelim
                g.FillEllipse(Brushes.Blue, depotEntrance.X, depotEntrance.Y, 20, 20);
                g.DrawString("Depo Girişi", new Font("Arial", 8), Brushes.Black, depotEntrance.X, depotEntrance.Y - 10);
            }
            pictureBoxMap.Image = map;
        }
    }

    public class Order
    {
        public int ProductId { get; set; }
        public string Corridor { get; set; }
        public int Rack { get; set; }
        public int Cell { get; set; }

        public Order(int productId, string corridor, int rack, int cell)
        {
            ProductId = productId;
            Corridor = corridor;
            Rack = rack;
            Cell = cell;
        }
    }
}
