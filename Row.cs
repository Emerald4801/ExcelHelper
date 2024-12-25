using ProjectManagement.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectManagement.Models
{
    internal class Row
    {
        public CellValue[] _Values { get; private set; }
        public int _Length { get; private set; }

        #region Constructors
        public Row(int Length) 
        {
            _Length = Length;
            _Values = new CellValue[Length];
        }

        public Row(CellValue[] Values)
        {
            _Length = Values.Length;
            _Values = Values;
        }

        #endregion

        #region Setters

        public void SetValues(CellValue[] Values) // ДОПИСАТЬ ПРОВЕРКУ НА ДЛИНУ
        {
            if (Values.Length != _Length)
            {
                return;
            }
            _Length = Values.Length;
            _Values = Values;
        }

        public void SetValue(int index, CellValue Value)
        {
            if (index >= _Length)
            {
                return;
            }
            _Values[index] = Value;
        }
        #endregion

        #region Getters

        public string GetValue(int index, out CellTypes type)
        {
            if(index >= _Length)
            {
                type = CellTypes.String;
                return "null";
            }

            type = _Values[index].Type;
            return _Values[index].Value;
        }

        public CellValue GetValue(int index)
        {
            if (index >= _Length)
            {
                return new CellValue();
            }
            return _Values[index];
        }
        #endregion

        #region Other

        public Row Concat(Row row)
        {
            //Создаем новую строку
            var newRow = new Row(this._Length + row._Length);
            //Записываем данные с первой
            for (int i = 0; i < this._Length; i++)
            {
                newRow._Values[i] = this._Values[i];
            }
            //Записываем данные со второй
            for (int i = this._Length, j = 0; i < newRow._Length; i++, j++)
            {
                newRow._Values[i] = row._Values[j];
            }
            return newRow;
        }

        #endregion
    }
}
