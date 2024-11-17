using System.Runtime.InteropServices;
using Ex = Microsoft.Office.Interop.Excel;
namespace ExTest
{
    internal class ExcelHelper : IDisposable
    {


        private Ex.Application _app;

        private Ex.Workbook _workbook;

        private Ex.Workbooks _workbooks;

        private Ex.Worksheet _worksheet;

        private string _path;

        public ExcelHelper()
        {
            _app = new Ex.Application();
        }
        public bool Open(string path, int idx)
        {
            try
            {
                _path = path;
                _workbooks = _app.Workbooks;
                if (!File.Exists(path))
                {
                    _workbook = _workbooks.Add();
                }
                else
                {
                    _workbook = _workbooks.Open(path);
                }

                _worksheet = _workbook.Worksheets[idx];

                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("ашибка" + e);
            }
            return false;
        }

        public string Get(int row, int col) 
        {
            return _worksheet.Cells[row, col].Value;
        }

        public void Set(int row, int col, string value)
        {
            _worksheet.Cells[row, col].Value = value;
        }

        //public string[] GetRow(int row)
        //{
            
        //}

        //public void SetRow(int row, int col, string value)
        //{
        //    _worksheet.Cells[row, col].Value = value;
        //}
        public void Save()
        {
            try
            {
                _workbook.SaveAs(_path);
            }
            catch (Exception e)
            {
                Console.WriteLine("ашибка " + e);
            }
        }
        public void Dispose()
        {
            //Освобождение _worksheet
            if (_worksheet != null)
            {
                while (Marshal.ReleaseComObject(_worksheet) != 0) { }
                _worksheet = null;
            }

            //Освобождение _workbook
            if (_workbook != null)
            {
                _workbook.Close();
                while (Marshal.ReleaseComObject(_workbook) != 0) { }
                _worksheet = null;
            }

            //Освобождение _workbooks
            if(_workbooks != null)
            {
                _workbooks.Close();
                while (Marshal.ReleaseComObject(_workbooks) != 0) { }
                _workbooks = null;
            }

            //Освобождение _app
            if(_app != null)
            {
                _app.Quit();
                while (Marshal.ReleaseComObject(_app) != 0) { }
                _app = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
