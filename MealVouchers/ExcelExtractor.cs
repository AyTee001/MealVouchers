using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MealVouchers
{
    public class ExcelExtractor : IDisposable
    {
        public ExcelExtractor(string fileName, string realSheetName)
        {
            _doc = SpreadsheetDocument.Open(fileName, false);
            _sheet = _doc.WorkbookPart.Workbook.Descendants<Sheet>().Where(x => x.Name == realSheetName).FirstOrDefault();
            _wsPart = (WorksheetPart)_doc.WorkbookPart.GetPartById(_sheet.Id);


            if (_doc == null || _sheet == null || _wsPart == null)
            {
                throw new InvalidDataException("Requested element is missing");
            }
        }

        public IReadOnlyList<IReadOnlyList<string>> GetStrings(IReadOnlyList<IReadOnlyList<string>> wantedCells)
        {
            var result = new List<IReadOnlyList<string>>();
            foreach (var cellNames in wantedCells)
            {
                var row = new List<string>();
                foreach (var cellName in cellNames)
                {
                    string value = GetCellValueString(cellName);
                    row.Add(value);
                }
                result.Add(row);
            }
            return result;
        }


        public string GetCellValueString(string addressName)
        {
            Cell theCell = _wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == addressName).FirstOrDefault();

            if (theCell == null)
            {
                throw new InvalidDataException($"Cell {addressName} not found");
            }
            var cellValue = theCell.CellValue.InnerText;
            if (theCell.DataType != null && theCell.DataType.Value == CellValues.SharedString)
            {
                var stringTable = _doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                cellValue = stringTable.SharedStringTable.ElementAt(int.Parse(cellValue)).InnerText;
            }
            return cellValue;
        }


        public bool CheckCell(string cellToCheck)
        {
            if (_sheet == null)
            {
                throw new InvalidDataException("One of the sheets is missimg or has an invalid name");
            }
            Cell theCell = _wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == cellToCheck).FirstOrDefault();
            if (theCell == null || theCell.CellValue == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }


        public DateTime GetDate(string dateCell)
        {
            Cell theCell = _wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == dateCell).FirstOrDefault();
            var cellValue = theCell.CellValue.InnerText;
            var newValue = DateTime.FromOADate(Convert.ToDouble(cellValue));
            return newValue;
        }


        public void Dispose()
        {
            _doc.Dispose();
        }


        #region data members
        private readonly SpreadsheetDocument _doc;
        private readonly Sheet? _sheet;
        private readonly WorksheetPart _wsPart;
        #endregion data members
    }

}
