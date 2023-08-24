using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Data;
using System.Reflection;
using System.Threading;

namespace WriteLargeExcelFileEfficiently
{

    public class DemoExcelCreator
    {
        public static bool CreateLargeExcel(DataSet ds, Stream xlsxmem)
        {
            //create a timer to figure out how long this takes
            Stopwatch watch = new Stopwatch();
            watch.Start();

            //event for thread tracking
            ManualResetEventSlim resetEvent = new ManualResetEventSlim(false);
            int toProcess = 0;

            Console.WriteLine(string.Format("{0}: Staring CreateLargeExcel", DateTime.Now.ToString()));

            //start our spreadsheet
            using (var spreadSheet = SpreadsheetDocument.Create(xlsxmem, SpreadsheetDocumentType.Workbook))
            {
                // create the workbook
                var workbookPart = spreadSheet.AddWorkbookPart();

                var openXmlExportHelper = new OpenXmlWriterHelper();
                openXmlExportHelper.SaveCustomStylesheet(workbookPart);

                var workbook = workbookPart.Workbook = new Workbook();
                var sheets = workbook.AppendChild<Sheets>(new Sheets());

                // Loop through each DataTable in DataSet and create worksheet for each
                uint worksheetNumber = 1;


                foreach (DataTable dt in ds.Tables)
                {
                    Console.WriteLine(string.Format("{0}: Staring WorkSheet {1}", DateTime.Now.ToString(), dt.TableName));

                    //setup the worksheet
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = worksheetNumber++, Name = dt.TableName };
                    sheets.Append(sheet);

                    toProcess++;
                    ThreadPool.QueueUserWorkItem(delegate
                    {
                        //open writer
                        using (var writer = OpenXmlWriter.Create(worksheetPart))
                        {

                            writer.WriteStartElement(new Worksheet());
                            writer.WriteStartElement(new SheetData());

                            //get number of columns
                            int numberOfColumns = dt.Columns.Count;
                            int numberofRows = dt.Rows.Count;
                            string[] ColDataType = new string[numberOfColumns];
                            decimal[] ColDistRowPerc = new decimal[numberOfColumns];
                            DateTime cellDateTime;

                            //here we are finding the number of distinct items in each column, helps to decide whether to used sharedsheet or not.
                            var ColsDistRowCount = dt.Columns
                                .Cast<DataColumn>()
                                .Select(dc => new
                                {
                                    colIndx = dc.Ordinal,
                                    Name = dc.ColumnName,
                                    Values = dt.Rows
                                        .Cast<DataRow>()
                                        .Select(row => row[dc])
                                        .Distinct()
                                        .Count()
                                })
                                .OrderBy(item => item.colIndx);

                            if (numberofRows > 0)
                            {
                                var ColsDistRowArray = ColsDistRowCount.Select(cols => cols.Values).ToArray();

                                for (int i = 0; i < numberOfColumns; i++)
                                {
                                    ColDistRowPerc[i] = Decimal.Multiply(Decimal.Divide(ColsDistRowArray[i], numberofRows), 100);
                                }
                            }


                            //Create header row
                            writer.WriteStartElement(new Row());
                            for (int i = 0; i < numberOfColumns; i++)
                            {
                                //header formatting attribute.  This will create a <c> element with s=2 as its attribute
                                //s stands for styleindex
                                DataColumn col = dt.Columns[i];
                                var attributes = new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, "2") }.ToList();
                                openXmlExportHelper.WriteCellValueSax(writer, col.ColumnName, CellValues.SharedString, attributes);

                                //get data type of column
                                ColDataType[i] = col.DataType.FullName;

                            }
                            writer.WriteEndElement(); //end of Row tag

                            foreach (DataRow dr in dt.Rows)
                            {
                                writer.WriteStartElement(new Row());

                                //write each column of data
                                for (int i = 0; i < numberOfColumns; i++)
                                {

                                    switch (ColDataType[i])
                                    {

                                        case "System.String":
                                            //For Strings, if distinct values are greater than percent here use inline, otherwise shared.
                                            if (ColDistRowPerc[i] >= 60)
                                            {
                                                openXmlExportHelper.WriteCellValueSax(writer, dr.ItemArray[i].ToString(), CellValues.InlineString);
                                            }
                                            else
                                            {
                                                openXmlExportHelper.WriteCellValueSax(writer, dr.ItemArray[i].ToString(), CellValues.SharedString);
                                            }
                                            break;
                                        case "System.DateTime":
                                            //date format.  Excel internally represent the datetime value as number, the date is only a formatting
                                            //applied to the number.  It will look something like 40000.2833 without formatting

                                            var attributes = new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, "3") }.ToList();

                                            //the helper internally translate the CellValues.Date into CellValues.Number before writing
                                            if (DateTime.TryParse(dr.ItemArray[i].ToString(), out cellDateTime))
                                            {
                                                if (cellDateTime.Ticks < 599317056000000000) //Excel reproduced a bug from Lotus for compatibility reasons, 29/02/1900 didn't actually exist.
                                                {

                                                    openXmlExportHelper.WriteCellValueSax(writer, (cellDateTime.ToOADate() - 1).ToString(CultureInfo.InvariantCulture), CellValues.Date, attributes);
                                                }
                                                else
                                                {
                                                    openXmlExportHelper.WriteCellValueSax(writer, cellDateTime.ToOADate().ToString(CultureInfo.InvariantCulture), CellValues.Date, attributes);
                                                }
                                            }
                                            else
                                            {
                                                openXmlExportHelper.WriteCellValueSax(writer, dr.ItemArray[i].ToString(), CellValues.Number, attributes);
                                            }
                                            break;
                                        default:
                                            //all other data types should fall under number format in excel
                                            openXmlExportHelper.WriteCellValueSax(writer, dr.ItemArray[i].ToString(), CellValues.Number);
                                            break;


                                    } // end switch coldatatype

                                } // end foreach column

                                writer.WriteEndElement(); // end row

                            } // end of foreach datarow


                            writer.WriteEndElement(); //end of SheetData
                            writer.WriteEndElement(); //end of worksheet
                            writer.Close();

                        } // using openxmlwrite

                        // If we're the last thread, signal
                        if (Interlocked.Decrement(ref toProcess) == 0)
                            resetEvent.Set();

                        Console.WriteLine(string.Format("{0}: Finished WorkSheet {1}", DateTime.Now.ToString(), dt.TableName));

                    });


                } // foreach datatable
                resetEvent.Wait();

                Console.WriteLine(string.Format("{0}: Starting CreateShareStringPart", DateTime.Now.ToString()));

                openXmlExportHelper.CreateShareStringPart(workbookPart);

            } // using spreadsheet
            watch.Stop();
            Console.WriteLine(string.Format("{0}: Finished CreateLargeExcel, Elapsed {1} ms", DateTime.Now.ToString(), watch.ElapsedMilliseconds));
            return true;
        } // createlargeexcel

    } //end class

} //end namespace