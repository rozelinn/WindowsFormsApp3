using OfficeOpenXml;
using System;
using System.Windows.Forms;
using WarehousePathOptimization;
using WindowsFormsApp3;

internal static class Program
{
    /// <summary>
    /// Uygulamanın ana girdi noktası.
    /// </summary>
    [STAThread]
    static void Main()
    {
        // ExcelPackage'i başlatmak için bir yol oluşturulabilir.
        // Burada örnek olarak bir excel dosyasından yükleme yapılması gerekir.
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // ExcelPackage dosyasını bir şekilde almanız gerekebilir (örneğin, bir dosyadan yükleyebilirsiniz).
        ExcelPackage excelPackage = new ExcelPackage(); // Burada uygun bir şekilde ExcelPackage'i başlatın.

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        // Form1'e excelPackage nesnesi geçirilir.
        Application.Run(new Form1(excelPackage));
    }
}
