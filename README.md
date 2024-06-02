# XML to PDF/CSV Converter

This repository contains a tool for converting XML files to either PDF or CSV format using C# and iTextSharp. The tool supports customization for different XML structures, including those exported from Excel.

## Features

- **XML to PDF Conversion**: Convert XML files to PDF format, preserving the structure and content.
- **XML to CSV Conversion**: Convert XML files to CSV format, with support for both generic XML and Excel-exported XML.
- **Attribute Handling**: Properly handles XML attributes and nested elements during conversion.
- **Image Embedding**: Supports embedding base64 encoded images within the XML into the PDF.
- **Customizable Output**: Provides utilities for customizing the output format, including CamelCase conversion.

## Prerequisites

- .NET 6.0 SDK or higher
- Visual Studio or any C# compatible IDE
- iTextSharp library

## Installation

1. **Clone the Repository**:
    ```bash
    git clone https://github.com/yogesh-nepal/XML-TO-CSV.git
    cd xml-to-pdf-csv
    ```

2. **Open the Project**:
   - Open the `XML2CSV.sln` solution file in Visual Studio or your preferred C# IDE.

## Usage

### Converting XML to PDF

1. **Specify the Input and Output Files**:
    - Update the `fileNamePath` and `outputFile` variables in the `Main` method of `Program.cs` with your XML file path and desired PDF output path.

2. **Run the Program**:
    - Build and run the project using your IDE. The program will generate a PDF file at the specified location.

### Converting XML to CSV

1. **Specify the Input and Output Files**:
    - Update the `fileNamePath` and `outputFile` variables in the `Main` method of `Program.cs` with your XML file path and desired CSV output path.
    - Choose whether the XML is exported from Excel by setting the `excelExported` parameter in the `ConvertToCSV` method.

2. **Run the Program**:
    - Build and run the project using your IDE. The program will generate a CSV file at the specified location.

## Example

### XML to PDF

**Input**: `1066214.xml`
**Output**: `TestData5.pdf`

```csharp
string fileNamePath = "D://[PROJECT]//XML2CSV//XML2CSV//XML//1066214.xml";
string outputFile = "D://[PROJECT]//XML2CSV//XML2CSV//Output//TestData5.pdf";

ConvertToPDF(fileNamePath, outputFile);
