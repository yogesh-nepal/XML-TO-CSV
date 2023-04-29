using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace XML2CSV
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string fileNamePath = "D://[PROJECT]//XML2CSV//XML2CSV//XML//1066214.xml";
            string outputFile = "D://[PROJECT]//XML2CSV//XML2CSV//Output//TestData5.pdf";

            ConvertToPDF(fileNamePath, outputFile);

            //Using XML exported from EXCEL
            //ConvertToCSV(fileNamePath, outputFile, true);

            //Using default XML 
            //ConvertToCSV(fileNamePath, outputFile, false);

        }

        private static void ConvertToPDF(string fileNamePath, string outputFile)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(fileNamePath);
            Document pdfDoc = new Document();
            PdfWriter.GetInstance(pdfDoc, new FileStream(outputFile, FileMode.Create));
            pdfDoc.Open();

            var nodes = doc.ChildNodes;
            List<XmlNode> filteredData = nodes.Cast<XmlNode>().ToList();
            var stuffToRemove = nodes.Cast<XmlNode>().ToList().Where(x => x.LocalName.ToLower().Contains("xml")).ToList();
            foreach (var item in stuffToRemove)
            {
                filteredData.Remove(item);
            }
            XDocument xdoc = XDocument.Load(fileNamePath);
            var root = xdoc.Root;
            var xElements = xdoc.Descendants().OfType<XElement>().FirstOrDefault();
            Font font = FontFactory.GetFont("Arial", 16, Font.NORMAL | Font.UNDERLINE | Font.BOLD, BaseColor.BLACK);
            Paragraph title = new Paragraph(GetStringFromLowerCamelCase(xElements.Name.LocalName), font);
            title.Alignment = Element.ALIGN_LEFT;
            pdfDoc.Add(title);
            Console.WriteLine(xElements.Name.LocalName);
            CreatePdf(xElements, pdfDoc);
            pdfDoc.Close();
            Console.ReadLine();
        }

        private static void CreatePdf(XElement xElements, Document pdfDoc)
        {
            Font font = FontFactory.GetFont("Arial", 12, Font.NORMAL, BaseColor.BLACK);

            if (xElements.Elements().ToList().Count() > 0)
            {
                foreach (XElement element in xElements.Elements().ToList())
                {
                    var hasChild = element.Elements().ToList().Count() > 0;
                    if (hasChild)
                    {
                        string text = element.Name.LocalName;
                        Console.WriteLine(text);
                        pdfDoc.Add(new Paragraph(" ", font));
                        pdfDoc.Add(new Paragraph(GetStringFromLowerCamelCase(text), font));
                        pdfDoc.Add(new Paragraph(" ", font));
                        var hasAttributes = element.HasAttributes;
                        if (hasAttributes)
                        {
                            string iText = "";
                            string textName = element.Name.LocalName;
                            var attributes = element.Attributes().ToList();
                            foreach (var item in attributes)
                            {
                                var attrValue = GetStringFromLowerCamelCase(item.Name.LocalName) + (!string.IsNullOrEmpty(item.Value) ? " : " + item.Value : string.Empty);
                                if (iText != "")
                                {
                                    iText += ", " + attrValue;
                                }
                                else
                                {
                                    iText = attrValue;
                                }
                            }
                            Font dataFont = FontFactory.GetFont("Arial", 12, Font.NORMAL, BaseColor.BLACK);
                            pdfDoc.Add(new Paragraph(GetStringFromLowerCamelCase(textName) + " | " + iText, dataFont));
                            Console.WriteLine(textName + " | " + iText);
                        }
                        //pdfDoc.Add(new Paragraph(text, font));
                        CreatePdf(element, pdfDoc);
                    }
                    else
                    {
                        var hasAttributes = element.HasAttributes;
                        string iText = "";
                        if (hasAttributes)
                        {
                            string text = element.Name.LocalName;
                            var attributes = element.Attributes().ToList();
                            foreach (var item in attributes)
                            {
                                var attrValue = GetStringFromLowerCamelCase(item.Name.LocalName) + (!string.IsNullOrEmpty(item.Value) ? " : " + item.Value : string.Empty);
                                if (iText != "")
                                {
                                    iText += ", " + attrValue;
                                }
                                else
                                {
                                    iText = attrValue;
                                }
                            }
                            pdfDoc.Add(new Paragraph(GetStringFromLowerCamelCase(text) + " | " + iText, font));
                            Console.WriteLine(text + " | " + iText);
                        }
                        else
                        {
                            string innerText = GetStringFromLowerCamelCase(element.Name.LocalName) + (!string.IsNullOrEmpty(element.Value) ? " : " + element.Value : string.Empty);
                            if (element.Value.Contains("data:image/png;base64"))
                            {
                                Regex regex = new Regex(@"^data:image/(?<mediaType>[^;]+);base64,(?<data>.*)");
                                Match match = regex.Match(element.Value);
                                Image image = Image.GetInstance(Convert.FromBase64String(match.Groups["data"].Value));
                                pdfDoc.Add(image);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(element.Value))
                                {
                                    var keepTogether = new PdfPTable(1)
                                    {
                                        KeepTogether = true,
                                        WidthPercentage = 100f
                                    };
                                    keepTogether.DefaultCell.Border = Rectangle.NO_BORDER;
                                    keepTogether.AddCell(new Paragraph(innerText, font));
                                    pdfDoc.Add(keepTogether);
                                    Console.WriteLine(innerText);
                                }
                            }
                        }
                    }
                }
            }
        }

        private static void ConvertToCSV(string fileNamePath, string outputFile, bool excelExported)
        {
            if (excelExported)
            {
                // Load the Excel XML file
                XmlDocument doc = new XmlDocument();
                doc.Load(fileNamePath);
                XmlNamespaceManager nsMgr = new XmlNamespaceManager(doc.NameTable);
                nsMgr.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet");

                XmlNodeList rowNodes = doc.SelectNodes("//ss:Table/ss:Row", nsMgr);

                // Set the path to the output CSV file
                // Create a new StreamWriter object to write to the output file
                using (var writer = new StreamWriter(outputFile))
                {
                    foreach (XmlNode row in rowNodes)
                    {
                        XmlNodeList cellNodes = row.ChildNodes;
                        List<XmlNode> cellList = cellNodes.Cast<XmlNode>().ToList();

                        string[] fields = cellList.Select(field => QuoteField(field.InnerText.ToString()))
                            .ToArray();
                        writer.WriteLine(string.Join(",", fields));
                    }
                }
            }
            else
            {
                // Create a new DataSet object
                DataSet ds = new();

                // Read the XML file into the DataSet
                ds.ReadXml(fileNamePath);
                DataTable dataTable = ds.Tables[0];
                outputFile = "D://[PROJECT]//XML2CSV//XML2CSV//Output//TestData1.csv";
                using (var writer = new StreamWriter(outputFile))
                {
                    // Write the header row
                    string[] headers = dataTable.Columns
                        .Cast<DataColumn>()
                        .Select(column => QuoteField(column.ColumnName))
                        .ToArray();
                    writer.WriteLine(string.Join(",", headers));

                    // Write the data rows
                    foreach (DataRow row in dataTable.Rows)
                    {
                        string[] fields = row.ItemArray
                            .Select(field => QuoteField(field.ToString()))
                            .ToArray();
                        writer.WriteLine(string.Join(",", fields));
                    }
                }
            }
        }

        // Encloses a field in double quotes if it contains a comma
        static string QuoteField(string field)
        {
            // If the field is null or empty, return an empty string
            if (string.IsNullOrEmpty(field)) return string.Empty;

            // If the field contains commas or newlines, enclose it in double quotes
            if (field.Contains(',') || field.Contains('\r') || field.Contains('\n'))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }

            // Otherwise, return the original field value
            return field;
        }
        static string GetStringFromLowerCamelCase(string input)
        {

            // Split the string into separate words using a regular expression
            string[] words = Regex.Split(input, @"(?<!^)(?=[A-Z])");

            // Capitalize the first letter of each word and join them together
            string output = string.Join(" ", words.Select(w => char.ToUpper(w[0]) + w.Substring(1)));

            return output;

        }
    }
}