using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace GenericExcelExport.Misc
{
    public class UtilExcel
    {
        /// <param name="filename">Name of the Excel file.</param>
        /// <param name="Source">The Generic Object List you want to export.</param>
        /// <param name="columnDefinition">This method was based on a grid export so you export just the columns you want and if you wish you can put a name diferent of the properties on the header.</param>
        /// <param name="reportInfo">Information of The Report.</param>
        /// <param name="reportInfo">Bool that determine if the sheet is going to be protected with a password.</param>
        public static bool ExportToExcel<V>(string filename, List<V> Source, Dictionary<string, string> columnDefinition, List<string> reportInfo, bool isProtected = false)
        {
            var path = System.IO.Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName;
            string extension = ".xlsx";
            string imageFilePath = path + ConfigurationManager.AppSettings["Logo"].ToString();

            string originalFileName = Path.GetFileNameWithoutExtension(filename) + extension;
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Reporte");
                worksheet.Cells.Style.Font.Size = 10;

                #region Cabecera
                int indiceLabel = 5;
                int indiceValor = 6;

                worksheet.Cells[2, indiceLabel].Value = "Usuario :";
                worksheet.Cells[2, indiceLabel].Style.Font.Bold = true;
                worksheet.Cells[3, indiceLabel].Value = "Compañia :";
                worksheet.Cells[3, indiceLabel].Style.Font.Bold = true;
                worksheet.Cells[4, indiceLabel].Value = "Fecha :";
                worksheet.Cells[4, indiceLabel].Style.Font.Bold = true;
                worksheet.Cells[5, indiceLabel].Value = "Reporte :";
                worksheet.Cells[5, indiceLabel].Style.Font.Bold = true;

                worksheet.Cells[2, indiceValor].Value = reportInfo[0];
                worksheet.Cells[3, indiceValor].Value = reportInfo[1];
                worksheet.Cells[4, indiceValor].Value = DateTime.Now.ToShortDateString() + " - " + DateTime.Now.ToShortTimeString();
                worksheet.Cells[5, indiceValor].Value = reportInfo[2];

                #endregion

                int index = 1;
                List<string> Columns = new List<string>();
                foreach (KeyValuePair<string, string> keyvalue in columnDefinition)
                {
                    worksheet.Cells[7, index].Value = keyvalue.Key;
                    index++;
                    Columns.Add(keyvalue.Value);
                }

                int row = 8;
                int col = 0;

                #region Llenar Data

                foreach (var dataItem in (System.Collections.IEnumerable)Source)
                {
                    col = 1;
                    foreach (string column in Columns)
                    {
                        foreach (PropertyInfo property in dataItem.GetType().GetProperties())
                        {
                            if (column.ToUpper() == property.Name.ToUpper())
                            {

                                if (property.PropertyType == typeof(Nullable<DateTime>) || property.PropertyType == typeof(DateTime))
                                {
                                    if (string.IsNullOrEmpty(System.Web.UI.DataBinder.GetPropertyValue(dataItem, property.Name, null)))
                                        worksheet.Cells[row, col].Value = string.Empty;
                                    else
                                    {
                                        worksheet.Cells[row, col].Style.Numberformat.Format = "@";
                                        //worksheet.Cells[row, col].Style.Numberformat.Format = "dd MMM yyyy hh:mm";
                                        worksheet.Cells[row, col].Value = System.Web.UI.DataBinder.GetPropertyValue(dataItem, property.Name, null);
                                    }
                                }
                                else if (property.PropertyType == typeof(Nullable<bool>) || property.PropertyType == typeof(bool))
                                {
                                    string value = System.Web.UI.DataBinder.GetPropertyValue(dataItem, property.Name, null);
                                    worksheet.Cells[row, col].Value = (string.IsNullOrEmpty(value) ? "" : (value == "True" ? "Si" : "No"));
                                }
                                else if (property.PropertyType == typeof(Nullable<Int16>) || property.PropertyType == typeof(Nullable<Int32>) || property.PropertyType == typeof(Nullable<Int64>))
                                {
                                    if (string.IsNullOrEmpty(System.Web.UI.DataBinder.GetPropertyValue(dataItem, property.Name, null)))
                                        worksheet.Cells[row, col].Value = string.Empty;
                                    else
                                        worksheet.Cells[row, col].Value = int.Parse(System.Web.UI.DataBinder.GetPropertyValue(dataItem, property.Name, null).ToString());
                                }
                                else if (property.PropertyType == typeof(Nullable<double>))
                                {
                                    if (string.IsNullOrEmpty(System.Web.UI.DataBinder.GetPropertyValue(dataItem, property.Name, null)))
                                        worksheet.Cells[row, col].Value = string.Empty;
                                    else
                                        worksheet.Cells[row, col].Value = double.Parse(System.Web.UI.DataBinder.GetPropertyValue(dataItem, property.Name, null).ToString());
                                }
                                else if (property.PropertyType == typeof(Nullable<decimal>))
                                {
                                    if (string.IsNullOrEmpty(System.Web.UI.DataBinder.GetPropertyValue(dataItem, property.Name, null)))
                                        worksheet.Cells[row, col].Value = string.Empty;
                                    else
                                        worksheet.Cells[row, col].Value = decimal.Parse(System.Web.UI.DataBinder.GetPropertyValue(dataItem, property.Name, null).ToString());
                                }
                                else if (property.PropertyType == typeof(Int16) || property.PropertyType == typeof(Int32) || property.PropertyType == typeof(Int64))
                                {
                                    worksheet.Cells[row, col].Style.Numberformat.Format = "#,##0";
                                    if (System.Web.UI.DataBinder.GetPropertyValue(dataItem, property.Name, null) == int.MinValue.ToString())
                                        worksheet.Cells[row, col].Value = string.Empty;
                                    else
                                        worksheet.Cells[row, col].Value = int.Parse(System.Web.UI.DataBinder.GetPropertyValue(dataItem, property.Name, null).ToString());
                                }
                                else if (property.PropertyType == typeof(decimal))
                                {
                                    worksheet.Cells[row, col].Style.Numberformat.Format = "#,##0.00";
                                    worksheet.Cells[row, col].Value = decimal.Parse(System.Web.UI.DataBinder.GetPropertyValue(dataItem, property.Name, null).ToString());
                                }
                                else if (property.PropertyType == typeof(double))
                                {
                                    worksheet.Cells[row, col].Style.Numberformat.Format = "#,##0.00";
                                    worksheet.Cells[row, col].Value = double.Parse(System.Web.UI.DataBinder.GetPropertyValue(dataItem, property.Name, null).ToString());
                                }
                                else
                                    worksheet.Cells[row, col].Value = System.Web.UI.DataBinder.GetPropertyValue(dataItem, property.Name, null);
                            }
                        }
                        col++;
                    }
                    row++;
                }

                #endregion

                using (var range = worksheet.Cells[7, 1, 7, Columns.Count()])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                    range.Style.Font.Color.SetColor(Color.White);
                    range.AutoFilter = true;
                }

                worksheet.Cells.AutoFitColumns(0);
                AddImageLogo(worksheet, 0, 1, imageFilePath);

                #region Proteccion por contraseña de la hoja de Excel

                if (isProtected)
                {
                    worksheet.Protection.SetPassword(ConfigurationManager.AppSettings["ExcelPassword"].ToString());
                    worksheet.Protection.AllowSelectLockedCells = true;
                    worksheet.Protection.AllowSort = true;
                    worksheet.Protection.AllowAutoFilter = true;
                    worksheet.Protection.AllowFormatCells = true;
                }

                #endregion
                using (var stream = new MemoryStream())
                    package.SaveAs(stream);

                #region Client Export

                Byte[] bin = package.GetAsByteArray();
                string file = filename + ".xlsx";
                File.WriteAllBytes(file, bin);

                //These lines will open it in Excel
                ProcessStartInfo pi = new ProcessStartInfo(file);
                Process.Start(pi);

                #endregion

                #region Web Export
                //HttpCookie cookieExport = new HttpCookie("excelExport");
                //cookieExport.Value = "1";

                //HttpContext.Current.Response.ClearHeaders();
                //HttpContext.Current.Response.Clear();
                //HttpContext.Current.Response.Buffer = false;
                //HttpContext.Current.Response.AddHeader("Content-disposition", "attachment; filename=" + originalFileName);
                //HttpContext.Current.Response.Charset = "UTF-8";
                //HttpContext.Current.Response.AppendCookie(cookieExport);
                //HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.Private);
                //HttpContext.Current.Response.ContentType = "application/octet-stream";
                //HttpContext.Current.Response.BinaryWrite(stream.ToArray());
                //HttpContext.Current.Response.Flush();
                //HttpContext.Current.Response.End();
                #endregion
            }
            return true;
        }

        private static void AddImageLogo(ExcelWorksheet ws, int columnIndex, int rowIndex, string filePath)
        {
            Bitmap image = new Bitmap(filePath);
            ExcelPicture picture = null;
            if (image != null)
            {
                picture = ws.Drawings.AddPicture("pic" + rowIndex.ToString() + columnIndex.ToString(), image);
                picture.From.Column = columnIndex;
                picture.From.Row = rowIndex;
                picture.From.ColumnOff = ConfigurarPixeles(2);
                picture.From.RowOff = ConfigurarPixeles(2);
                picture.SetSize(220, 70);
            }
        }

        public static int ConfigurarPixeles(int pixels)
        {
            int mtus = pixels * 9525;
            return mtus;
        }

    }
}
