using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;

using System.Reflection;


namespace ObjectsListToExcel
{
    public class ExcelGenrator
    {
        public ExcelGenrator() { }
        public Worksheet addHeader(Workbook wb, Worksheet ws, List<Object> objs, string title, string fontFamily = "Sakkal Majalla", string color = "#3498DB")
        {
            // Sheet initialisation.
            // var ws = wb.Worksheets.Add("nomDeLaListe").SetTabColor("#3498DB");

            // font choice.
            // ws.Style.Font.FontName = fontFamily;
            //ws.Style.Font.SetFontSize(13);
            ////ws.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //ws.Style.Alignment.WrapText = true;
            Object obj = objs.FirstOrDefault();
            // Add the model fields to the header of the excel file.
            int totalOfFields = obj.GetType().GetProperties().Length; // number of fields in the object.
            int numberOfFields = 0;
            // Adding the title of table in excel file.
            // ws.get_Range(ws.Cells[4, 4], ws.Cells[4, totalOfFields + 3]).Merge().Value = title;

            Microsoft.Office.Interop.Excel.Range chartRangeTitle;
            Range c1 = (Range)ws.Cells[4, 4];
            Range c2 = (Range)ws.Cells[4, totalOfFields + 3];
            chartRangeTitle = (Range)ws.get_Range(c1, c2);
            chartRangeTitle.Merge();
            chartRangeTitle.Value = title;
            chartRangeTitle.Interior.Color = Color.CornflowerBlue;
           // chartRangeTitle.Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
           // chartRangeTitle.Style.VerticalAlignment = XlHAlign.xlHAlignCenter;

            // ws.get_Range(ws.Cells[4, 4], ws.Cells[4, totalOfFields + 3]).Merge().Style.Fill.BackgroundColor = Color.Red;
            //// ws.Range(ws.Cell(4, 4), ws.Cell(4, totaffflOfFields + 3)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            // ws.get_Range(ws.Cells[4, 4], ws.Cells[4, totalOfFields + 3]).Style.Font.Bold = true;
            // ws.get_Range(ws.Cells[4, 4], ws.Cells[4, totalOfFields + 3]).Style.Font.FontColor = Color.Red;
            // ws.get_Range(ws.Cells[4, 4], ws.Cells[4, totalOfFields + 3]).Style.Font.FontSize = 18;
            // Looping all propeties of the object.
            foreach (var prop in obj.GetType().GetProperties())
            {
                var displayNameAttribute = prop.GetCustomAttributes(typeof(DisplayNameAttribute), false);
                string displayName = prop.Name;
                if (displayNameAttribute.Count() != 0)
                {
                    displayName = (displayNameAttribute[0] as DisplayNameAttribute).DisplayName;
                }
                numberOfFields++;
                ws.Cells[5, totalOfFields - numberOfFields + 4] = displayName;

              
               // ws.Columns[totalOfFields - numberOfFields + 4].Style.Font.Bold = true;
            }

            return ws;
        }
        public Worksheet addBody(Worksheet ws, List<Object> objs)
        {
            int numberOfFields = 0;
            int numberOfRecords = 0;
            Object obj = objs.FirstOrDefault();
            int totalOfFields = obj.GetType().GetProperties().Length;
            string previousValue = "";
            int indexOfPreviousValue = 0;

            foreach (var item in objs.ToList())
            {
                numberOfFields = 0;
                Type myType = item.GetType();
                IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());

                foreach (PropertyInfo prop in props)
                {
                    object propValue = prop.GetValue(item, null);

                    numberOfFields++;

                    ws.Cells[6 + numberOfRecords, totalOfFields - numberOfFields + 4] = propValue;
                  //  ws.Cells[6 + numberOfRecords, totalOfFields - numberOfFields + 4].Style.Font.Bold = true;

                    //ws.Cells[6 + numberOfRecords, totalOfFields - numberOfFields + 4].Border.
                    //ws.Cells[6 + numberOfRecords, totalOfFields - numberOfFields + 4].Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    //ws.Cells[6 + numberOfRecords, totalOfFields - numberOfFields + 4].Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    //ws.Cells[6 + numberOfRecords, totalOfFields - numberOfFields + 4].Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    if (numberOfFields == 1 && numberOfRecords == 0)
                    {
                        previousValue = propValue.ToString();
                    }
                    else
                    {
                        if (numberOfFields == 1)
                        {
                            if (previousValue == propValue.ToString())
                            {
                                //ws.get_Range(ws.Cells[6 + numberOfRecords - (1 + indexOfPreviousValue), totalOfFields - numberOfFields + 4], ws.Cells[6 + numberOfRecords, totalOfFields - numberOfFields + 4]).Merge(true).Value = propValue.ToString(); 
                                //ws.get_Range(ws.Cells[6 + numberOfRecords - (1 + indexOfPreviousValue), totalOfFields - numberOfFields + 4], ws.Cells[6 + numberOfRecords, totalOfFields - numberOfFields + 4]).Merge(true).Value = propValue.ToString();

                                Microsoft.Office.Interop.Excel.Range chartRangeTitle;
                                Range c11 = (Range)ws.Cells[6 + numberOfRecords - (1 + indexOfPreviousValue), totalOfFields - numberOfFields + 4];
                                Range c22 = (Range)ws.Cells[6 + numberOfRecords, totalOfFields - numberOfFields + 4];
                                c22.Value = "";
                                chartRangeTitle = (Range)ws.get_Range(c11, c22);
                                chartRangeTitle.Merge();
                                chartRangeTitle.Value = propValue.ToString();

                                //         ws.Range(ws.Cell(6 + numberOfRecords - (1 + indexOfPreviousValue), totalOfFields - numberOfFields + 4), ws.Cell(6 + numberOfRecords, totalOfFields - numberOfFields + 4)).Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                //        ws.Range(ws.Cell(6 + numberOfRecords - (1 + indexOfPreviousValue), totalOfFields - numberOfFields + 4), ws.Cell(6 + numberOfRecords, totalOfFields - numberOfFields + 4)).Merge().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                indexOfPreviousValue++;
                            }
                            else
                            {
                                previousValue = propValue.ToString();
                                indexOfPreviousValue = 0;
                            }
                        }

                    }


                }
                numberOfRecords++;
            }
            Range chartRange;
            Range c1 = (Range)ws.Cells[5, 4];
            Range c2 = (Range)ws.Cells[5 + numberOfRecords, numberOfFields + 3];
            chartRange = (Range)ws.get_Range(c1, c2);
            chartRange.EntireColumn.AutoFit();
            foreach (Range cell in chartRange.Cells)
            {
                cell.BorderAround2();
            }
            return ws;
        }
        public void Generate1()
        {
            Application ExcelApp = new Application();
            Workbook ExcelWorkBook = null;
            Worksheet ExcelWorkSheet = null;
            ExcelApp.Visible = true;
            ExcelWorkBook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            // ExcelWorkBook.Worksheets.Add(); //Adding New Sheet in Excel Workbook
            try
            {
                ExcelWorkSheet = (Worksheet)ExcelWorkBook.Worksheets[1]; // Compulsory Line in which sheet you want to write data
                //Writing data into excel of 100 rows with 10 column 
                for (int r = 1; r < 101; r++) //r stands for ExcelRow and c for ExcelColumn
                {
                    // Excel row and column start positions for writing Row=1 and Col=1
                    for (int c = 1; c < 11; c++)
                        ExcelWorkSheet.Cells[r, c] = "R" + r + "C" + c;
                }

              //  ExcelWorkBook.Worksheets[1].Name = "MySheet";//Renaming the Sheet1 to MySheet

                ExcelWorkBook.SaveAs("C://Users//dell//Desktop//LIB//ABCD.xlsx");

                ExcelWorkBook.Close();

                ExcelApp.Quit();



            }

            catch (Exception exHandle)

            {

                Console.WriteLine("Exception: " + exHandle.Message);

                Console.ReadLine();

            }
        }
        public void Generate(List<Object> objs, string title, string fontFamily = "Sakkal Majalla", string color = "#3498DB")
        {
            Application ExcelApp = new Application();
            Workbook wb = null;
            Worksheet ws = null;
            ExcelApp.Visible = true;
            wb = ExcelApp.Workbooks.Add();
            ws = (Worksheet)wb.Worksheets[1];
            ws = addHeader(wb, ws, objs, title);
            ws = addBody(ws, objs);
            string path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\TestExcelGen20.xlsx";
            if (System.IO.File.Exists(path))
            {
                System.IO.File.Delete(path);
            }
            wb.SaveAs(path);

            //wb.Worksheets[1].Name = "Liste";//Renaming the Sheet1 to MySheet


            wb.Close();

            ExcelApp.Quit();

            //byte[] fileBytes = System.IO.File.ReadAllBytes(tempFile);
            //if (System.IO.File.Exists(tempFile))
            //{
            //    System.IO.File.Delete(tempFile);
            //}
            

           // wc.DownloadFile(tempFile, "test322222.xlsx");
            //System.Web.HttpContext.Current.Response.Clear();
            // Response.Buffer = true;
            //HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=ttest688.xlsx");
            //HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
            //HttpContext.Current.Response.BinaryWrite(fileBytes);
            //HttpContext.Current.Response.End();

        }

    }
}
