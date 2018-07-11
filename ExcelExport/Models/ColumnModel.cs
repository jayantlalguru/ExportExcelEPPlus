using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport.Models
{
    /// <summary>
    /// Column name and its width. Other property of column can be added here if required
    /// </summary>
    public class ColumnModel
    {
        public ColumnModel()
        {
            ColumnWidth = 25;
        }
        /// <summary>
        /// Name of the column
        /// </summary>
        public string ColumnName { get; set; }
        /// <summary>
        /// Width of the column
        /// </summary>
        public int ColumnWidth { get; set; }
    }
}
