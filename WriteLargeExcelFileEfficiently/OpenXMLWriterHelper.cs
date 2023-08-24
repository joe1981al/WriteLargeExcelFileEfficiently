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
    /// <summary>
    /// A helper class to keep track of the sharedstringtable string being created in a spreadsheet
    /// </summary>
    public class OpenXmlWriterHelper
    //class extracted from this project: http://www.codeproject.com/Articles/877791/How-to-Create-Large-Excel-File-using-Openxml
    {
        //lock to make thread safe
        ReaderWriterLockSlim SharedDictLock = new ReaderWriterLockSlim();

        /// <summary>
        /// contains the shared string as the key, and the index as the value.  index is 0 base
        /// </summary>
        private readonly Dictionary<string, int> _shareStringDictionary = new Dictionary<string, int>();
        private int _shareStringMaxIndex = 0;

        /// <summary>
        /// create the default excel formats.  These formats are required for the excel in order for it to render
        /// correctly.
        /// </summary>
        /// <returns></returns>
        private Stylesheet CreateDefaultStylesheet()
        {

            Stylesheet ss = new Stylesheet();

            Fonts fts = new Fonts();
            DocumentFormat.OpenXml.Spreadsheet.Font ft = new DocumentFormat.OpenXml.Spreadsheet.Font();
            FontName ftn = new FontName();
            ftn.Val = "Calibri";
            FontSize ftsz = new FontSize();
            ftsz.Val = 11;
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            fts.Append(ft);
            fts.Count = (uint)fts.ChildElements.Count;

            Fills fills = new Fills();
            Fill fill;
            PatternFill patternFill;

            //default fills used by Excel, don't changes these

            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.None;
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);

            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Gray125;
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);



            fills.Count = (uint)fills.ChildElements.Count;

            Borders borders = new Borders();
            Border border = new Border();
            border.LeftBorder = new LeftBorder();
            border.RightBorder = new RightBorder();
            border.TopBorder = new TopBorder();
            border.BottomBorder = new BottomBorder();
            border.DiagonalBorder = new DiagonalBorder();
            borders.Append(border);
            borders.Count = (uint)borders.ChildElements.Count;

            CellStyleFormats csfs = new CellStyleFormats();
            CellFormat cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            csfs.Append(cf);
            csfs.Count = (uint)csfs.ChildElements.Count;


            CellFormats cfs = new CellFormats();

            cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cfs.Append(cf);



            var nfs = new NumberingFormats();



            nfs.Count = (uint)nfs.ChildElements.Count;
            cfs.Count = (uint)cfs.ChildElements.Count;

            ss.Append(nfs);
            ss.Append(fts);
            ss.Append(fills);
            ss.Append(borders);
            ss.Append(csfs);
            ss.Append(cfs);

            CellStyles css = new CellStyles(
                new CellStyle()
                {
                    Name = "Normal",
                    FormatId = 0,
                    BuiltinId = 0,
                }
                );

            css.Count = (uint)css.ChildElements.Count;
            ss.Append(css);

            DifferentialFormats dfs = new DifferentialFormats();
            dfs.Count = 0;
            ss.Append(dfs);

            TableStyles tss = new TableStyles();
            tss.Count = 0;
            tss.DefaultTableStyle = "TableStyleMedium9";
            tss.DefaultPivotStyle = "PivotStyleLight16";
            ss.Append(tss);
            return ss;
        }


        virtual public void SaveCustomStylesheet(WorkbookPart workbookPart)
        {

            //get a copy of the default excel style sheet then add additional styles to it
            var stylesheet = CreateDefaultStylesheet();

            // ***************************** Fills *********************************
            var fills = stylesheet.Fills;

            //header fills background color
            var fill = new Fill();
            var patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Solid;
            patternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("C8EEFF") };
            //patternFill.BackgroundColor = new BackgroundColor() { Indexed = 64 };
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);
            fills.Count = (uint)fills.ChildElements.Count;

            // *************************** numbering formats ***********************
            var nfs = stylesheet.NumberingFormats;
            //number less than 164 is reserved by excel for default formats
            uint iExcelIndex = 165;
            NumberingFormat nf;
            nf = new NumberingFormat();
            nf.NumberFormatId = iExcelIndex++;
            nf.FormatCode = @"[$-409]m/d/yy\ h:mm\ AM/PM;@";
            nfs.Append(nf);

            nfs.Count = (uint)nfs.ChildElements.Count;

            //************************** cell formats ***********************************
            var cfs = stylesheet.CellFormats;//this should already contain a default StyleIndex of 0

            var cf = new CellFormat();// Date time format is defined as StyleIndex = 1
            cf.NumberFormatId = nf.NumberFormatId;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = true;
            cfs.Append(cf);

            cf = new CellFormat();// Header format is defined as StyleINdex = 2
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 2;
            cf.ApplyFill = true;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cfs.Append(cf);

            cf = new CellFormat();// Date time format as defined in StyleIndex = 3
            cf.NumberFormatId = 14;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = true;
            cfs.Append(cf);


            cfs.Count = (uint)cfs.ChildElements.Count;

            var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            var style = workbookStylesPart.Stylesheet = stylesheet;
            style.Save();

        }


        /// <summary>
        /// write out the share string xml.  Call this after writing out all shared string values in sheet
        /// </summary>
        /// <param name="workbookPart"></param>
        public void CreateShareStringPart(WorkbookPart workbookPart)
        {
            if (_shareStringMaxIndex > 0)
            {
                var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                using (var writer = OpenXmlWriter.Create(sharedStringPart))
                {
                    writer.WriteStartElement(new SharedStringTable());
                    foreach (var item in _shareStringDictionary)
                    {
                        writer.WriteStartElement(new SharedStringItem());
                        writer.WriteElement(new Text(item.Key));
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                }
            }

        }


        /// <summary>
        /// CellValues = Boolean -> expects cellValue "True" or "False"
        /// CellValues = InlineString -> stores string within sheet
        /// CellValues = SharedString -> stores index within sheet. If this is called, please call CreateShareStringPart after creating all sheet data to create the shared string part
        /// CellValues = Date -> expects ((DateTime)value).ToOADate().ToString(CultureInfo.InvariantCulture) as cellValue 
        ///              and new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, "1") }.ToList() as attributes so that the correct formatting can be applied
        /// 
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="cellValue"></param>
        /// <param name="dataType"></param>
        /// <param name="attributes"></param>
        public void WriteCellValueSax(OpenXmlWriter writer, string cellValue, CellValues dataType, List<OpenXmlAttribute> attributes = null)
        {
            switch (dataType)
            {
                case CellValues.InlineString:
                    {
                        if (attributes == null)
                        {
                            attributes = new List<OpenXmlAttribute>();
                        }
                        attributes.Add(new OpenXmlAttribute("t", null, "inlineStr"));
                        writer.WriteStartElement(new Cell(), attributes);
                        writer.WriteElement(new InlineString(new Text(cellValue)));
                        writer.WriteEndElement();
                        break;
                    }
                case CellValues.SharedString:
                    //locks added in this section to work with threading
                    {
                        if (attributes == null)
                        {
                            attributes = new List<OpenXmlAttribute>();
                        }
                        attributes.Add(new OpenXmlAttribute("t", null, "s"));//shared string type
                        writer.WriteStartElement(new Cell(), attributes);
                        SharedDictLock.EnterUpgradeableReadLock();
                        {
                            if (!_shareStringDictionary.ContainsKey(cellValue))
                            {
                                SharedDictLock.EnterWriteLock();
                                _shareStringDictionary.Add(cellValue, _shareStringMaxIndex);
                                _shareStringMaxIndex++;
                                SharedDictLock.ExitWriteLock();
                            }
                        }

                        //writing the index as the cell value

                        writer.WriteElement(new CellValue(_shareStringDictionary[cellValue].ToString()));
                        SharedDictLock.ExitUpgradeableReadLock();


                        writer.WriteEndElement();//cell

                        break;
                    }
                case CellValues.Date:
                    {
                        if (attributes == null)
                        {
                            writer.WriteStartElement(new Cell() { DataType = CellValues.Number });
                        }
                        else
                        {
                            writer.WriteStartElement(new Cell() { DataType = CellValues.Number }, attributes);
                        }

                        writer.WriteElement(new CellValue(cellValue));

                        writer.WriteEndElement();

                        break;
                    }
                case CellValues.Boolean:
                    {
                        if (attributes == null)
                        {
                            attributes = new List<OpenXmlAttribute>();
                        }
                        attributes.Add(new OpenXmlAttribute("t", null, "b"));//boolean type
                        writer.WriteStartElement(new Cell(), attributes);
                        writer.WriteElement(new CellValue(cellValue == "True" ? "1" : "0"));
                        writer.WriteEndElement();
                        break;
                    }
                default:
                    {
                        if (attributes == null)
                        {
                            writer.WriteStartElement(new Cell() { DataType = dataType });
                        }
                        else
                        {
                            writer.WriteStartElement(new Cell() { DataType = dataType }, attributes);
                        }
                        writer.WriteElement(new CellValue(cellValue));

                        writer.WriteEndElement();


                        break;
                    }
            }

        } // WriteCellValueSax
    } // end class
} //end namespace