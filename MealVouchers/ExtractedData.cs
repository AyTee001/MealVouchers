using System.Globalization;

namespace MealVouchers
{
    public class ExtractedData
    {
        public ExtractedData(string date, string personalNumber, IReadOnlyList<IReadOnlyList<string>> values)
        {
            Date = date;
            EmployeeNumber = personalNumber;
            ReferenceValueAccountNumber = values[0][0];
            TaxFreeAccountNumber = values[1][0];
            ReferenceValue = decimal.Parse(values[0][1], CultureInfo.InvariantCulture);
            TaxFree = decimal.Parse(values[1][1], CultureInfo.InvariantCulture);
        }

        public void PrintExtractedData(TextWriter writer)
        {
            writer.WriteLine("{0};{1};{2};{3}", 
                Date, EmployeeNumber, 
                ReferenceValueAccountNumber,
                Math.Round(ReferenceValue, 2).ToString(CultureInfo.InvariantCulture));
            writer.WriteLine("{0};{1};{2};{3}",
                Date, EmployeeNumber,
                TaxFreeAccountNumber,
                Math.Round(TaxFree, 2).ToString(CultureInfo.InvariantCulture));
            writer.WriteLine();
            writer.WriteLine();
        }

        public string Date { get; }
        public string EmployeeNumber { get; }
        public string  ReferenceValueAccountNumber { get; }
        public string TaxFreeAccountNumber { get; }
        public decimal ReferenceValue { get; }
        public decimal TaxFree { get; }
    }
}
