using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport.Models
{
    /// <summary>
    /// Model for excel workbook
    /// </summary>
    public class WorkbookModel
    {
        public WorkbookModel()
        {
            FileName = Guid.NewGuid().ToString() + ".xlsx";
            WorksheetModels = new List<WorksheetModel>();
            ImageModel = new ImageModel();            
        }
        /// <summary>
        /// Total number of worksheet
        /// </summary>
        public List<WorksheetModel> WorksheetModels { get; set; }
        /// <summary>
        /// Name of the excel file
        /// </summary>
        public string FileName { get; set; }
        /// <summary>
        /// Directory where the excel file will be stored
        /// </summary>
        public string FilePath { get; set; }
        /// <summary>
        /// Properties of image
        /// </summary>
        public ImageModel ImageModel { get; set; }
        /// <summary>
        /// Disclaimer to be displayed in all worksheets
        /// </summary>
        public string Disclaimer { get; set; }
    }
}
