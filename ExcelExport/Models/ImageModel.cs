using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport.Models
{
    /// <summary>
    /// Set image properties
    /// </summary>
    public class ImageModel
    {   
        public ImageModel()
        {
            ColumnIndex = 0;
            RowIndex = 0;
        }
        /// <summary>
        /// Start column index for image
        /// </summary>
        public int ColumnIndex { get; set; }
        /// <summary>
        /// Start row index for image
        /// </summary>
        public int RowIndex { get; set; }
        /// <summary>
        /// Complete image path with image name Ex: C:\Users\Downloads\SomeImage.jpg
        /// </summary>
        public string ImagePath { get; set; }
    }
}
