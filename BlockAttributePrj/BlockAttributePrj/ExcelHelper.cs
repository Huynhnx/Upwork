using Autodesk.AutoCAD.Geometry;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace BlockAttributePrj
{
    class ExcelHelper
    {
        /// <summary>
        /// Get string value of an Excel Cell.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = string.Empty;
            if (cell.CellValue != null)
            {
                value = cell.CellValue.InnerXml;
            }
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }
        public static bool IsFileLocked(string filepath)
        {
            FileInfo file = new FileInfo(filepath);
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
        public static Point3d ToPoint3d(string s)
        {
            string[] a = s.Trim('(', ')').Split(',');
            return new Point3d(double.Parse(a[0]), double.Parse(a[1]), double.Parse(a[2]));
        }
    }
}
