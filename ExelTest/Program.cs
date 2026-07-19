using System;
using System.Diagnostics;
using ClosedXML.Excel;
using System.Windows.Forms; 
using System.Drawing;
using System.Security.Policy;
using Path = System.IO.Path;

namespace ExcelTest
{
    internal class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            ExcelHandler excelHandler = new ExcelHandler();
            
            string columnQR = "R";
            string columnToChange = "V";
            excelHandler.ChangePaymentType(columnQR, columnToChange);
            
            HashSet<string> necessaryColumns = ["g", "o", "p", "v"];
            excelHandler.Copy(necessaryColumns);

            string columnToSortBy = "C";
            excelHandler.Sort(columnToSortBy);
            
            Dictionary<string, string> paymentTypes= new()
            {
                {"Cash", "E"},
                {"Card", "F"},
                {"qr", "G"}
            };
            string columnToCheckType = "D";
            string columnToGetMoney = "A";
            excelHandler.FillCellBasedOn(paymentTypes, columnToCheckType, columnToGetMoney);

            Dictionary<string, XLColor> services = new()
            {
                {"50.01 - 62.01 Оплата за курьерские услуги", XLColor.Yellow},
                {"50.01 - 76.10.1 Поступление НП от клиента", XLColor.Red},
            };
            string columnToCheck = "C";
            excelHandler.PaintCells(services, columnToCheck);

            HashSet<string> columns = ["D", "C"];
            excelHandler.Delete(columns);
            
            excelHandler.SaveWorkbook();
        }
    }
    
    public class ExcelHandler
    {
        private string _oldExcelExtension = ".xls";
        private string _newExcelExtension = ".xlsx";
        private string _filter = "Excel файлы (*.xlsx;*.xls)|*.xlsx;*.xls|Все файлы (*.*)|*.*";
        private int _sheetIndex = 1;
        private int _firstRowIndex = 1;
        
        private XLWorkbook _workbook;
        private IXLWorksheet _firstWorksheet;
        private IXLWorksheet _secondWorksheet;

        public ExcelHandler()
        {
            _workbook = GetWorkbook();
            _firstWorksheet = _workbook.Worksheet(_sheetIndex);
            _firstWorksheet.Row(_firstRowIndex).Delete();
        }
        
        public void Sort(string columnToSortBy)
        {
            var rows = _secondWorksheet.RowsUsed().Count();
            
            var range = _secondWorksheet.Range($"A1:D{rows}");

            range.Sort(columnToSortBy);
        }

        public void ChangePaymentType(string columnQr, string secondColumn)
        {
            var columnQR = _firstWorksheet.Column(columnQr);
            
            var columnToChange = _firstWorksheet.Column(secondColumn);

            var cells = columnQR.CellsUsed();

            foreach (var cell in cells)
            {
                string value = cell.Value.GetText();
                
                if (value != string.Empty)
                {
                    var row = cell.WorksheetRow().RowNumber();

                    columnToChange.Cell(row).Value = value;
                }
            }

            columnQR.Hide();
        }
        
        public void FillCellBasedOn(Dictionary<string, string> paymentTypes, string type, string money)
        {
            var rows = _secondWorksheet.RowsUsed();
            var columnMoney = _secondWorksheet.Column(money);
            var columnType = _secondWorksheet.Column(type);
                
            foreach (var (paymentType, paymentColumn) in paymentTypes)
            {
                var columnToFill = _secondWorksheet.Column(paymentColumn);

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
        
        public void Copy(HashSet<string> necessaryColumns)
        {
            _secondWorksheet = _workbook.AddWorksheet();

            int count = 1;
            
            foreach (var column in necessaryColumns)
            {
                _firstWorksheet.Column(column).CopyTo(_secondWorksheet.Column(count));

                count++;
            }
            
            _firstWorksheet.Delete();
        }
        
        public void PaintCells(Dictionary<string, XLColor> services, string columnToCheck)
        {
            var columnCheck = _secondWorksheet.Column(columnToCheck);
            var rows = _secondWorksheet.RowsUsed();
            
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
        
        public void Delete(HashSet<string> columns)
        {
            foreach (var name in columns)
            {
                _secondWorksheet.Column(name).Delete();
            }
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