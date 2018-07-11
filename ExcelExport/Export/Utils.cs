using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelExport.Export
{
    public class Utils
    {
        static DirectoryInfo _outputDir = null;
        public static DirectoryInfo OutputDir
        {
            get
            {
                return _outputDir;
            }
            set
            {
                _outputDir = value;
                if (!_outputDir.Exists)
                {
                    _outputDir.Create();
                }
            }
        }
        public static FileInfo GetFileInfo(string file, string filePath, bool deleteIfExists = true)
        {
            //var fi = new FileInfo(OutputDir.FullName + Path.DirectorySeparatorChar + file);
            var fi = new FileInfo(filePath + file);
            if (deleteIfExists && fi.Exists)
            {
                fi.Delete();  // ensures we create a new workbook
            }
            return fi;
        }
        public static FileInfo GetFileInfo(DirectoryInfo altOutputDir, string file, bool deleteIfExists = true)
        {
            var fi = new FileInfo(altOutputDir.FullName + Path.DirectorySeparatorChar + file);
            if (deleteIfExists && fi.Exists)
            {
                fi.Delete();  // ensures we create a new workbook
            }
            return fi;
        }

        internal static DirectoryInfo GetDirectoryInfo(string directory)
        {
            var di = new DirectoryInfo(_outputDir.FullName + Path.DirectorySeparatorChar + directory);
            if (!di.Exists)
            {
                di.Create();
            }
            return di;
        }
        /// <summary>
        /// If columnName = 1 then return value is A. If columnName = 2 then return value is B
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        internal static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
        /// <summary>
        /// If name = A the return value = 1. If name = B the return value = 2
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static int GetColumnNumber(string name)
        {
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
        }

        /// <summary>
        /// returns only numeric value from a string parameter
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static int GetNumberFromString(string value)
        {   
            string str = string.Empty;
            int val = 0;

            for (int i = 0; i < value.Length; i++)
            {
                if (Char.IsDigit(value[i]))
                    str += value[i];
            }

            if (str.Length > 0)
                val = int.Parse(str);
            return val;
        }
    }
}
