using ClosedXML.Excel;
using Path = System.IO.Path;

namespace ExelTest
{
    internal class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            int additionalWorksheets = 1;

            ExcelHandler excelHandler = new ExcelHandler(additionalWorksheets);

            if(!excelHandler.IsCreated)
                return;

            int sheetIndex = 1;
            string columnQr = "R";
            string secondColumn = "V";
            excelHandler.ChangePaymentType(columnQr, secondColumn, sheetIndex);

            HashSet<string> necessaryColumns = ["g", "o", "p", "v"];
            int sheetIndexCopyTo = 2;
            excelHandler.Copy(necessaryColumns, sheetIndex,sheetIndexCopyTo);

            string columnToSortBy = "C";
            string startCell = "A1";
            excelHandler.Sort(columnToSortBy, startCell, sheetIndexCopyTo);

            Dictionary<string, string> paymentTypes= new()
            {
                {"Cash", "E"},
                {"Card", "F"},
                {"qr", "G"}
            };
            string columnToCheckType = "D";
            string columnToGetMoney = "A";
            excelHandler.FillCellBasedOn(paymentTypes, columnToCheckType, columnToGetMoney, sheetIndexCopyTo);

            Dictionary<string, XLColor> services = new()
            {
                {"50.01 - 62.01 Оплата за курьерские услуги", XLColor.Yellow},
                {"50.01 - 76.10.1 Поступление НП от клиента", XLColor.Red},
            };
            string columnToCheck = "C";
            excelHandler.PaintCells(services, columnToCheck, sheetIndexCopyTo);

            HashSet<string> columns = ["D", "C"];
            excelHandler.Delete(columns, sheetIndexCopyTo);

            excelHandler.SaveWorkbook();
        }
    }

    public class ExcelHandler
    {
        private readonly string _oldExcelExtension = ".xls";
        private readonly string _newExcelExtension = ".xlsx";
        private readonly string _filter = "Excel файлы (*.xlsx;*.xls)|*.xlsx;*.xls|Все файлы (*.*)|*.*";

        private int _firstRowIndex = 1;

        private XLWorkbook _workbook;
        private IXLWorksheet? _currentWorksheet;

        public bool IsCreated => _workbook != null;

        public ExcelHandler(int additionalWorksheetsAmount)
        {
            _workbook = GetWorkbook();

            for (int i = 0; i < additionalWorksheetsAmount; i++)
            {
                _workbook.AddWorksheet();
            }

            var initialWorksheet = _workbook.Worksheets.First();

            initialWorksheet.Row(_firstRowIndex).Delete();
        }

        public void Sort(string columnToSortBy, string startCell, int sheetIndex)
        {
            _currentWorksheet = _workbook.Worksheet(sheetIndex);

            var rowsCount = _currentWorksheet.RowsUsed().Count();

            var column = _currentWorksheet.LastColumnUsed()?.Cell(rowsCount).WorksheetColumn().ColumnLetter();

            var range = _currentWorksheet.Range($"{startCell}:{column}{rowsCount}");

            range.Sort(columnToSortBy);
        }

        public void ChangePaymentType(string columnQr, string secondColumn, int sheetIndex)
        {
            _currentWorksheet = _workbook.Worksheet(sheetIndex);

            var firstColumn = _currentWorksheet.Column(columnQr);

            var columnToChange = _currentWorksheet.Column(secondColumn);

            var cells = firstColumn.CellsUsed();

            foreach (var cell in cells)
            {
                string value = cell.Value.GetText();

                if (value != string.Empty)
                {
                    var row = cell.WorksheetRow().RowNumber();

                    columnToChange.Cell(row).Value = value;
                }
            }

            firstColumn.Hide();
        }

        public void FillCellBasedOn(Dictionary<string, string> paymentTypes, string type, string money, int sheetIndex)
        {
            _currentWorksheet = _workbook.Worksheet(sheetIndex);

            var rows = _currentWorksheet.RowsUsed();
            var columnMoney = _currentWorksheet.Column(money);
            var columnType = _currentWorksheet.Column(type);

            foreach (var (paymentType, paymentColumn) in paymentTypes)
            {
                var columnToFill = _currentWorksheet.Column(paymentColumn);

                foreach (var row in rows)
                {
                    var position = row.RowNumber();

                    var cellWithType = columnType.Cell(position);

                    var verifier = cellWithType.GetText();

                    if (string.Equals(paymentType, verifier, StringComparison.CurrentCultureIgnoreCase))
                    {
                        columnToFill.Cell(position).Value = columnMoney.Cell(position).Value;
                    }
                }
            }
        }

        public void Copy(HashSet<string> necessaryColumns, int fromSheet, int toSheet)
        {
            _currentWorksheet = _workbook.Worksheet(fromSheet);
            var secondSheet = _workbook.Worksheet(toSheet);

            int count = 1;

            foreach (var column in necessaryColumns)
            {
                _currentWorksheet.Column(column).CopyTo(secondSheet.Column(count));

                count++;
            }
        }

        public void PaintCells(Dictionary<string, XLColor> services, string columnToCheck, int sheetIndex)
        {
            _currentWorksheet = _workbook.Worksheet(sheetIndex);

            var columnCheck = _currentWorksheet.Column(columnToCheck);
            var rows = _currentWorksheet.RowsUsed();

            foreach (var (service, color) in services)
            {
                foreach (var row in rows)
                {
                    var position = row.RowNumber();

                    var cellText = columnCheck.Cell(position).GetText();
                    var cell = columnCheck.Cell(position);

                    if (string.Equals(cellText, service, StringComparison.CurrentCultureIgnoreCase))
                    {
                        cell.CellLeft().Style.Fill.SetBackgroundColor(color);
                    }
                }
            }
        }

        public void Delete(HashSet<string> columns, int sheetIndex)
        {
            _currentWorksheet = _workbook.Worksheet(sheetIndex);

            foreach (var name in columns)
            {
                _currentWorksheet.Column(name).Delete();
            }
        }

        public void Delete(int sheetIndex)
        {
            _workbook.Worksheet(sheetIndex).Delete();
        }

        public void SaveWorkbook()
        {
            SaveFileDialog fileDialog = new SaveFileDialog();

            fileDialog.Filter = _filter;
            fileDialog.FilterIndex = 1;

            fileDialog.ShowDialog();

            _workbook.SaveAs(Path.GetFullPath(fileDialog.FileName));
        }

        private XLWorkbook GetWorkbook()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = _filter;
            openFileDialog.FilterIndex = 1;
            openFileDialog.Title = "Выберите Excel файл";

            var result = openFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                string extension = Path.GetExtension(openFileDialog.FileName);

                if (extension == _oldExcelExtension || extension == _newExcelExtension)
                {
                    _workbook = new XLWorkbook(openFileDialog.FileName);
                }

                return _workbook;
            }

            return null!;
        }
    }
}
