using ProjectManagement.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectManagement.Models
{
    internal struct CellValue
    {
        public CellValue() 
        {

        }

        public CellValue(string value, CellTypes type)
        {
            Value = value;
            Type = type;
        }

        public string Value { get; set; }
        public CellTypes Type { get; set; }
    }


}
