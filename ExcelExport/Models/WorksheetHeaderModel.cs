using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace ExcelExport.Models
{
    /// <summary>
    /// A model which holds the property of worksheet Header
    /// </summary>
    public class WorksheetHeaderModel
    {
        public WorksheetHeaderModel()
        {
            HeaderText = "Report Generated";
            HeaderFontFamily = "Arial Black";
            HeaderTextColor = Color.White;
            HeaderCell = "A11";
            HeaderCellRange = "A11:P11";
            MergeHeader = true;
            HeaderFontSize = 20;
            HeaderBackgroundColor = Color.DarkBlue;
            HorizontalAlignment = HorizontalAlignment.CenterContinuous;
        }
        /// <summary>
        /// Header Text for the worksheet
        /// </summary>
        public string HeaderText { get; set; }
        /// <summary>
        /// Font Family for header
        /// </summary>
        public string HeaderFontFamily { get; set; }
        /// <summary>
        /// Font Color of Header text
        /// </summary>
        public Color HeaderTextColor { get; set; }
        /// <summary>
        /// Header Cell Value where you want to put the header text Ex: A1
        /// </summary>
        public string HeaderCell { get; set; }
        /// <summary>
        /// Coloumn span range for header Ex: A1:G1
        /// </summary>
        public string HeaderCellRange { get; set; }
        /// <summary>
        /// If Header Cells needs to be merged
        /// </summary>
        public bool MergeHeader { get; set; }
        /// <summary>
        /// Font size of the header
        /// </summary>
        public int HeaderFontSize { get; set; }
        /// <summary>
        /// Bacground color for header
        /// </summary>
        public Color HeaderBackgroundColor { get; set; }
        /// <summary>
        /// Horizontal alignment property
        /// </summary>
        public HorizontalAlignment HorizontalAlignment { get; set; }
        /// <summary>
        /// Vertical alignment property
        /// </summary>
        public VerticalAlignment VerticalAlignment { get; set; }
    }
}
