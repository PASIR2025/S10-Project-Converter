using ClosedXML.Excel;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
namespace S10ProjectConverter;
public partial class MainWindow : Window
{
    private readonly ObservableCollection<BudgetItem> _items = new();
    public MainWindow() { InitializeComponent(); ItemsGrid.ItemsSource = _items; }
    private void Import_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx" };
        if (dialog.ShowDialog() != true) return;
        try
        {
            using var workbook = new XLWorkbook(dialog.FileName);
            var sheet = workbook.Worksheet(1);
            var rows = sheet.RangeUsed()?.RowsUsed().ToList() ?? new();
            if (rows.Count < 2) throw new InvalidDataException("El archivo no contiene filas suficientes.");
            // Formato inicial: Código | Descripción | Unidad | Metrado | Precio unitario
            var imported = new List<BudgetItem>();
            foreach (var row in rows.Skip(1))
            {
                var description = row.Cell(2).GetString().Trim();
                if (string.IsNullOrWhiteSpace(description)) continue;
                if (!TryNumber(row.Cell(4), out var quantity) || !TryNumber(row.Cell(5), out var price)) continue;
                imported.Add(new BudgetItem(row.Cell(1).GetString(), description, row.Cell(3).GetString(), quantity, price));
            }
            if (imported.Count == 0) throw new InvalidDataException("No se encontraron partidas numéricas. Se espera: Código, Descripción, Unidad, Metrado, Precio unitario.");
            _items.Clear();
            foreach (var item in imported) _items.Add(item);
            CountText.Text = _items.Count.ToString();
            TotalText.Text = _items.Sum(x => x.Subtotal).ToString("C2", CultureInfo.GetCultureInfo("es-PE"));
        }
        catch (Exception ex) { MessageBox.Show(ex.Message, "No se pudo importar", MessageBoxButton.OK, MessageBoxImage.Warning); }
    }
    private static bool TryNumber(IXLCell cell, out decimal result)
    {
        if (cell.TryGetValue<decimal>(out result)) return true;
        var value = cell.GetFormattedString().Trim();
        return decimal.TryParse(value, NumberStyles.Number, CultureInfo.GetCultureInfo("es-PE"), out result)
            || decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out result);
    }
    private void Export_Click(object sender, RoutedEventArgs e)
    {
        if (_items.Count == 0) { MessageBox.Show("Primero importa un presupuesto."); return; }
        var dialog = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv", FileName = "partidas.csv" };
        if (dialog.ShowDialog() != true) return;
        static string Esc(string s) => "\"" + s.Replace("\"", "\"\"") + "\"";
        var lines = new List<string> { "Codigo;Descripcion;Unidad;Metrado;PrecioUnitario;Parcial" };
        lines.AddRange(_items.Select(x => string.Join(";", Esc(x.Code), Esc(x.Description), Esc(x.Unit), x.Quantity.ToString(CultureInfo.InvariantCulture), x.UnitPrice.ToString(CultureInfo.InvariantCulture), x.Subtotal.ToString(CultureInfo.InvariantCulture))));
        File.WriteAllLines(dialog.FileName, lines, new UTF8Encoding(true));
        MessageBox.Show("Archivo CSV exportado correctamente.");
    }
}
public record BudgetItem(string Code, string Description, string Unit, decimal Quantity, decimal UnitPrice)
{
    public decimal Subtotal => Quantity * UnitPrice;
}