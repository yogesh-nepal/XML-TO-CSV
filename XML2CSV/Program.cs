using System.Data;

namespace XML2CSV
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Set the path to the XML file
            var fileNamePath = "yourPathToXMLFile";

            // Create a new DataSet object
            DataSet ds = new();

            // Read the XML file into the DataSet
            ds.ReadXml(fileNamePath);

            // Convert the first table in the DataSet to CSV and write it to a file
            ConvertToCSV(ds.Tables[0]);
        }

        private static void ConvertToCSV(DataTable dataTable)
        {
            // Set the path to the output CSV file
            var outputFile = "yourPathToExportCSVFile";

            // Create a new StreamWriter object to write to the output file
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

        // Encloses a field in double quotes if it contains a comma
        static string QuoteField(string field)
        {
            // If the field is null or empty, return an empty string
            if (string.IsNullOrEmpty(field)) return string.Empty;

            // If the field contains commas, enclose it in double quotes
            if (field.Contains(','))
            {
                return $"\"{field}\"";
            }

            // Otherwise, return the original field value
            return field;
        }
    }
}