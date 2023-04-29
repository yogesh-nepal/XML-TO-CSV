using System.Data;
using System.Xml;

namespace XML2CSV
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string fileNamePath = "D://[PROJECT]//XML2CSV//XML2CSV//XML//TestData.xml";
            string outputFile = "D://[PROJECT]//XML2CSV//XML2CSV//Output//TestData.csv";

            //Using XML exported from EXCEL
            ConvertToCSV(fileNamePath, outputFile, true);

            //Using default XML 
            ConvertToCSV(fileNamePath, outputFile, false);
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
    }
}