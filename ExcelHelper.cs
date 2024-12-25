
using Microsoft.Office.Interop.Excel;
using ProjectManagement.Helpers;
using ProjectManagement.Models;
using System.Drawing;
using System.Runtime.InteropServices;
using Ex = Microsoft.Office.Interop.Excel;


namespace InfrastructureFiles
{
    internal class ExcelHelper : IDisposable
    {
        private Ex.Application _app;
        private Ex.Workbook _workbook;
        private string _filePath;
        private Ex.Worksheet active;
        public ExcelHelper()
        {
            _app = new Ex.Application();
        }

        public void Dispose()
        {
            try
            {
                //Освобождение _worksheet
                if (active != null)
                {
                    while (Marshal.ReleaseComObject(active) != 0) { }
                    active = null;
                }

                //Освобождение _workbook
                if (_workbook != null)
                {
                    _workbook.Close();
                    while (Marshal.ReleaseComObject(_workbook) != 0) { }
                    _workbook = null;
                }

                ////Освобождение _workbooks
                //if (_workbooks != null)
                //{
                //    _workbooks.Close();
                //    while (Marshal.ReleaseComObject(_workbooks) != 0) { }
                //    _workbooks = null;
                //}

                //Освобождение _app
                if (_app != null)
                {
                    _app.Quit();
                    while (Marshal.ReleaseComObject(_app) != 0) { }
                    _app = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception e)
            {

            }
        }
        internal void OpenWorkSheet(int listIndex = 1)
        {
            active = _workbook.Worksheets[listIndex] as Worksheet;
            active.Activate();
        }
        internal bool Open(string filePath, int listIndex = 1)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    _workbook = _app.Workbooks.Open(filePath);
                    active = _workbook.Worksheets[listIndex] as Worksheet;
                    active.Activate();
                    _filePath = filePath;
                }
                else
                {
                    _workbook = _app.Workbooks.Add();
                    active = _workbook.Worksheets.Add();
                    active.Activate();
                    _filePath = filePath;
                }


                return true;
            }
            catch (Exception e)
            {

            }
            return false;
        }

        internal void Save()
        {
            if (!string.IsNullOrEmpty(this._filePath))
            {
                if (File.Exists(_filePath))
                {
                    _workbook.Save();
                }
                else
                {
                    _workbook.SaveAs(_filePath);
                }
                _filePath = null;
            }
        }
        internal void SaveAs(string Path)
        {
            if (!string.IsNullOrEmpty(Path))
            {
                if (File.Exists(Path))
                {
                    _workbook.Save();
                }
                else
                {
                    _workbook.SaveAs(Path);
                }
                //_filePath = Path;
            }
        }

        #region Setters
        internal void SetNameOfSheet(string nameOfSheet)
        {
            active.Name = nameOfSheet;
        }

        internal bool Set(int row, int column, object data, string type = "string")
        {
            try
            {
                Dictionary<string, string> TypesOfValue = new()
                {
                    {"string", "@" }, //Текстовый
                    {"double", "# ##0,00"}, //Числовой
                    {"integer", "# ##0"}, //Целочисленный
                    {"money", "# ##0,00р."} //Денежный
                };

                active.Cells[row, column].NumberFormat = TypesOfValue[type];             
                ((Ex.Worksheet)_app.ActiveSheet).Cells[row, column] = data;
                return true;
            }
            catch (Exception e)
            {

            }
            return false;
        }

        internal bool SetRow(int rowIdx, Row row, int StartCol = 1)
        {
            try
            {
                for (int i = StartCol, j = 0; i < StartCol + row._Length; j++, i++)
                {
                    Set(rowIdx + 1, i, row.GetValue(j).Value);
                }
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return false;
        }


        #endregion

        #region Getters

        internal string Get(int row, int column)
        {
            try
            {
                if (row < 1 || column < 1)
                {
                    Console.WriteLine("Номера строк и колонок не могут быть меньше единицы");
                    return "";
                }

                Ex.Range range = (Ex.Range)active.Cells[row, column];

                if (((Ex.Range)active.Cells[row, column]).Value is double)
                {
                    return Math.Round(double.Parse(((Ex.Range)active.Cells[row, column]).Text), 4).ToString();
                }

                if (!string.IsNullOrEmpty(((Ex.Range)active.Cells[row, column]).Text))
                {
                    return ((Ex.Range)active.Cells[row, column]).Text;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            return "";
        }

        internal Row GetRow(int rowIndex, int StartCol, int EndCol) //Дописать проверку на заполнение
        {
            var row = new Row(EndCol - StartCol + 1);

            for (int i = StartCol; i < EndCol; i++)
            {
                row.SetValue(i - StartCol, new CellValue(Get(rowIndex, i), CellTypes.String));
            }
            return row;
        }

        #endregion

        #region Design

        internal void Bold(int startRow, int startCol, int endCol, int endRow = -1)
        {
            if (endRow == -1)
                endRow = startRow;

            for (int i = startRow; i <= endRow; i++)
            {
                for (int j = startCol; j <= endCol; j++)
                    active.Cells[i, j].Font.Bold = true;
            }
        }



        internal bool Merge(string firstColumn, int firstRow, string secondColumn, int secondRow)
        {
            try
            {
                string Cell1 = $"{firstColumn}{firstRow}";
                string Cell2 = $"{secondColumn}{secondRow}";
                ((Ex.Worksheet)_app.ActiveSheet).Range[Cell1, Cell2].Merge(Type.Missing);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return false;
        }
        internal bool Merge(string Cell1, string Cell2)
        {
            try
            {
                ((Ex.Worksheet)_app.ActiveSheet).Range[Cell1, Cell2].Merge();//Много вопросов
                return true;
            }
            catch (Exception e)
            {

            }
            return false;
        }
        internal bool AutoFit(string StartColumn, string EndColumn)
        {
            try
            {
                //((Excel.Worksheet)_excel.ActiveSheet).Range[$"{StartColumn}1:{EndColumn}1"].AutoFit;
                return true;
            }
            catch (Exception e)
            {

            }
            return false;
        }
        internal void Paint(int row, int column, Color color)
        {
            ((Ex.Range)active.Cells[row, column]).Interior.Color = color;
        }

        internal void BorderLine(int startRow, int startCol, int endCol, int endRow = -1)
        {
            if (endRow == -1)
                endRow = startRow;

            for (int i = startRow; i <= endRow; i++)
            {
                for (int j = startCol; j <= endCol; j++)
                { 
                    var range = (Ex.Range)active.Cells[i, j];
                    range.BorderAround2();
                }
            }
        }
        #endregion

        public void CreateDropDownList(int row, int column, List<string> variants)
        {
            var flatList = string.Join(",", variants.ToArray());

            var cell = (Ex.Range)active.Cells[row, column];
            cell.Validation.Delete();
            cell.Validation.Add(
               XlDVType.xlValidateList,
               XlDVAlertStyle.xlValidAlertInformation,
               XlFormatConditionOperator.xlBetween,
               flatList,
               Type.Missing);

            cell.Validation.IgnoreBlank = true;
            cell.Validation.InCellDropdown = true;
        }

        /// <summary>
        /// Проверяет ячейку на факт объеденения
        /// </summary>
        /// <returns>Возвращает bool</returns>
        public bool CheckMerge(int row, int column)
        {
            return active.Cells[row, column].MergeCells;
        }
    }
}
