using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Data;

namespace ExcelExport.Models
{
    /// <summary>
    /// Model which holds properties of excel worksheet
    /// </summary>
    public class WorksheetModel
    {
        private static int SheetNumber;
        public WorksheetModel()
        {
            SheetNumber = SheetNumber + 1;
            WorksheetName = "Sheet" + SheetNumber.ToString();
            //WorksheetHeaderModel worksheetHeaderModel = new WorksheetHeaderModel();
            WorksheetHeaderModels = new List<WorksheetHeaderModel>();
            //WorksheetHeaderModels.Add(worksheetHeaderModel);
            DataColumnHeaderModel = new DataColumnHeaderModel();
            WorksheetDataStartRow = 13;
        }
        /// <summary>
        /// Name of the worksheet
        /// </summary>
        public string WorksheetName { get; set; }
        /// <summary>
        /// Header of worksheet
        /// </summary>
        public List<WorksheetHeaderModel> WorksheetHeaderModels { get; set; }
        /// <summary>
        /// Column Header
        /// </summary>
        public DataColumnHeaderModel DataColumnHeaderModel { get; set; }
        /// <summary>
        /// Data for the worksheet
        /// </summary>
        public DataTable WorksheetData { get; set; }
        /// <summary>
        /// Form which row data should start writing
        /// </summary>
        public int WorksheetDataStartRow { get; set; }
    }
}
