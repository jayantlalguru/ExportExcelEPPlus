using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport.Models
{
    /// <summary>
    /// A model which holds the property of column header
    /// </summary>
    public class DataColumnHeaderModel
    {
        public DataColumnHeaderModel()
        {
            Columns = new List<ColumnModel>();
            ColumnStartCell = "A";
            BackgroundColor = Color.LightBlue;
            Bold = true;
            StartRow = 14;
            FontColor = Color.Black;
        }
        /// <summary>
        /// Column header names to be displayed
        /// </summary>
        public List<ColumnModel> Columns { get; set; }
        /// <summary>
        /// From which cell the column header should start Ex: A5
        /// </summary>
        public string ColumnStartCell { get; set; }
        /// <summary>
        /// Backgroung colour of Header
        /// </summary>
        public Color BackgroundColor { get; set; }
        /// <summary>
        /// Column header bold property
        /// </summary>
        public bool Bold { get; set; }
        /// <summary>
        /// From which row the actual data needs to be displayed
        /// </summary>
        public int StartRow { get; set; }
        /// <summary>
        /// Color of the column header text
        /// </summary>
        public Color FontColor { get; set; }
        //public ExcelFillStyle MyProperty { get; set; }
    }
}
