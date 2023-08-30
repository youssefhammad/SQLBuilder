using OfficeOpenXml;
using System.Data.SqlClient;
using System.Text;

namespace SQLBuilder
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string excelFilePath = "C:\\Users\\MyLab\\Downloads\\Healthcare Provider Taxonomy.xlsx";
            string sheetName = "nucc_taxonomy_231";
            string tableName = "dbo.HealthcareProviderTaxonomies";
            string outputFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "result.txt");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set the license context

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];
                int rowCount = worksheet.Dimension.Rows;
                int columnCount = worksheet.Dimension.Columns;

                int conceptNameColumnIndex = 0;

                // Find the index of the "Concept Name" column
                for (int column = 1; column <= columnCount; column++)
                {
                    var columnName = worksheet.Cells[1, column].Value;
                    if (columnName != null && columnName.ToString() == "Concept Name")
                    {
                        conceptNameColumnIndex = column;
                        break;
                    }
                }

                if (conceptNameColumnIndex == 0)
                {
                    Console.WriteLine("Concept Name column not found.");
                    return;
                }

                List<string> insertStatements = new List<string>();

                for (int row = 2; row <= rowCount; row++)
                {
                    var conceptName = worksheet.Cells[row, conceptNameColumnIndex].Value;
                    if (conceptName == null || string.IsNullOrWhiteSpace(conceptName.ToString()))
                    {
                        continue;
                    }

                    string insertStatement = GenerateInsertStatement(tableName, worksheet, row, columnCount);
                    insertStatements.Add(insertStatement);
                }

                // Create a StringBuilder to store the INSERT statements
                StringBuilder sb = new StringBuilder();

                // Append the insert statements to the StringBuilder
                foreach (string statement in insertStatements)
                {
                    sb.AppendLine(statement);
                }

                // Create the directory if it doesn't exist
                Directory.CreateDirectory(Path.GetDirectoryName(outputFilePath));

                // Save the INSERT statements to the output file
                File.WriteAllText(outputFilePath, sb.ToString());
            }

            Console.WriteLine("Data import completed.");
        }

        static string GenerateInsertStatement(string tableName, ExcelWorksheet worksheet, int row, int columnCount)
        {
            string columns = "";
            string values = "";

            for (int column = 1; column <= columnCount; column++)
            {
                var cellValue = worksheet.Cells[1, column].Value;
                string columnName = cellValue != null ? cellValue.ToString() : string.Empty;

                // If the column is not Concept Code or Concept Name, skip it
                if (columnName != "Concept Code" && columnName != "Concept Name")
                {
                    continue;
                }

                cellValue = worksheet.Cells[row, column].Value;
                string cellData = cellValue != null ? cellValue.ToString() : string.Empty;

                // Remove spaces from the column name
                string formattedColumnName = columnName.Replace(" ", "");

                if (!string.IsNullOrEmpty(formattedColumnName) && !string.IsNullOrEmpty(cellData))
                {
                    columns += $"{formattedColumnName}, ";
                    values += $"'{cellData.Replace("'", "''")}', ";
                }
            }

            // Check if both columns and values are empty, skip generating the INSERT statement
            if (string.IsNullOrEmpty(columns) && string.IsNullOrEmpty(values))
            {
                return string.Empty;
            }

            // Remove the trailing comma and space
            columns = columns.TrimEnd(',', ' ');
            values = values.TrimEnd(',', ' ');

            return $"INSERT INTO {tableName} ({columns}) VALUES ({values})";
        }
    }
}
