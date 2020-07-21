[![Build Status](https://dev.azure.com/simplify9/Github%20Pipelines/_apis/build/status/simplify9.ExportToExcel?branchName=master)](https://dev.azure.com/simplify9/Github%20Pipelines/_build/latest?definitionId=168&branchName=master) 

![Azure DevOps tests](https://img.shields.io/azure-devops/tests/Simplify9/Github%20Pipelines/168?style=for-the-badge)


| **Package**       | **Version** |
| :----------------:|:----------------------:|
|```SimplyWorks.ExportToExcel```| ![Nuget](https://img.shields.io/nuget/v/SimplyWorks.ExportToExcel?style=for-the-badge)



## Introduction 
*ExportToExcel* is a library that provides extensions to [IEnumerable Interface](https://docs.microsoft.com/en-us/dotnet/api/system.collections.ienumerable?view=netcore-3.1) extensions. 

## Getting Started
*ExportToExcel* is available as a package on [NuGet](https://www.nuget.org/packages/SimplyWorks.ExportToExcel/). 

To use *ExportToExcel*, you will require the [`Documentformat.OpenXml`](https://www.nuget.org/packages/DocumentFormat.OpenXml/) package. 

## Functions Available 
1. *ExportToExcel*
```csharp
async public static Task<byte[]> ExportToExcel<TEntity>(this IEnumerable<TEntity> data)
       {
           var dictionary = typeof(TEntity).GetProperties().ToDictionary(k => k.Name, v => v.Name);
           return await ExportToExcel(data, dictionary);
       }

       async public static Task<byte[]> ExportToExcel<TEntity>(this IEnumerable<TEntity> data, IEnumerable<string> columns)
       {
           var dictionary = columns.ToDictionary(k => k, v => v);
           return await ExportToExcel(data, dictionary);
       }

       async public static Task<byte[]> ExportToExcel<TEntity>(this IEnumerable<TEntity> data, IDictionary<string, string> columns)
       {
           var tempFile = Path.GetTempFileName();

           try
           {
               await WriteExcel(data, tempFile, columns);

               return File.ReadAllBytes(tempFile);
           }
           finally
           {
               File.Delete(tempFile);
           }
       }
```
2. *WriteExcel*
```csharp 
async public static Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, string filePath)
{
    var dictionary = typeof(TEntity).GetProperties().ToDictionary(k => k.Name, v => v.Name);
    await WriteExcel(data, filePath, dictionary);
}

async public static Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, string filePath, IEnumerable<string> columns)
{
    var dictionary = columns.ToDictionary(k => k, v => v);
    await WriteExcel(data, filePath, dictionary);
}

async public static Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, string filePath, IDictionary<string, string> columns)
{
    using (var fileStream = File.Open(filePath, FileMode.Create, FileAccess.ReadWrite))
    {
        await WriteExcel(data, fileStream, columns);
    }
}

async public static Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, Stream stream)
{
    var dictionary = typeof(TEntity).GetProperties().ToDictionary(k => k.Name, v => v.Name);
    await WriteExcel(data, stream, dictionary);
}

async public static Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, Stream stream, IEnumerable<string> columns)
{
    var dictionary = columns.ToDictionary(k => k, v => v);
    await WriteExcel(data, stream, dictionary);
}

async public static Task WriteExcel<TEntity>(this IEnumerable<TEntity> data, Stream stream, IDictionary<string, string> columns)
{
    using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
    {
        var workbookPart = doc.AddWorkbookPart();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var worksheetPartId = workbookPart.GetIdOfPart(worksheetPart);
        var stylesheetPart = workbookPart.AddNewPart<WorkbookStylesPart>();

        // create the string table
        var sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
        var xmlStringTable = new XElement(mainNamespace + "sst");
        await WriteXmlToPartAsync(sharedStringTablePart, xmlStringTable);

        XElement workbookElement = new XElement(mainNamespace + "workbook",
            new XAttribute("xmlns", mainNamespace.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "r", relationshipsNamespace.NamespaceName),
            new XElement("bookViews",
            new XElement("workbookView")),
            new XElement("sheets",
            new XElement("sheet",
                new XAttribute("name", "Exported"),
                new XAttribute("sheetId", "1"),
                new XAttribute(relationshipsNamespace + "id", worksheetPartId))));

        foreach (var element in workbookElement.Descendants())

            if (element.Name.Namespace == "")
            {
                element.Attributes("xmlns").Remove();
                element.Name = element.Parent.Name.Namespace + element.Name.LocalName;
            }

        await WriteXmlToPartAsync(workbookPart, workbookElement);

        var propertyInfos = new List<PropertyInfo>();
        foreach (var propertyInfo in typeof(TEntity).GetProperties())

            switch (true)
            {
                case object _ when propertyInfo.PropertyType == typeof(string):
                case object _ when propertyInfo.PropertyType == typeof(DateTime):
                case object _ when propertyInfo.PropertyType == typeof(DateTime?):
                case object _ when propertyInfo.PropertyType == typeof(int):
                case object _ when propertyInfo.PropertyType == typeof(int?):
                case object _ when propertyInfo.PropertyType == typeof(byte):
                case object _ when propertyInfo.PropertyType == typeof(byte?):
                case object _ when propertyInfo.PropertyType == typeof(short):
                case object _ when propertyInfo.PropertyType == typeof(short?):
                case object _ when propertyInfo.PropertyType == typeof(long):
                case object _ when propertyInfo.PropertyType == typeof(long?):
                case object _ when propertyInfo.PropertyType == typeof(float):
                case object _ when propertyInfo.PropertyType == typeof(float?):
                case object _ when propertyInfo.PropertyType == typeof(double):
                case object _ when propertyInfo.PropertyType == typeof(double?):
                case object _ when propertyInfo.PropertyType == typeof(bool):
                case object _ when propertyInfo.PropertyType == typeof(bool?):
                case object _ when propertyInfo.PropertyType == typeof(decimal):
                case object _ when propertyInfo.PropertyType == typeof(decimal?):
                    if (columns.ContainsKey(propertyInfo.Name))
                        propertyInfos.Add(propertyInfo);
                    break;
            }

        XElement worksheetElement = new XElement(mainNamespace + "worksheet",
            new XElement("sheetViews",
                new XElement("sheetView",
                    new XAttribute("tabSelected", "1"),
                    new XAttribute("workbookViewId", "0")),
                new XElement("pane",
                    new XAttribute("ySplit", "1"),
                    new XAttribute("topLeftCell", "A2"),
                    new XAttribute("activePane", "bottomLeft"),
                    new XAttribute("state", "frozen")),
                new XElement("selection",
                    new XAttribute("pane", "bottomLeft"))),
            new XElement("sheetFormatPr",
                new XAttribute("defaultRowHeight", "15")),
                new XElement("cols", GetCols(propertyInfos)),
            new XElement("sheetData", GetRowValues(data, columns, propertyInfos)));

        foreach (var element in worksheetElement.Descendants())

            if (element.Name.Namespace == "")
            {
                element.Attributes("xmlns").Remove();
                element.Name = element.Parent.Name.Namespace + element.Name.LocalName;
            }

        await WriteXmlToPartAsync(worksheetPart, worksheetElement);

        var styleSheetElement = XElement.Parse(styleSheet);

        await WriteXmlToPartAsync(stylesheetPart, styleSheetElement);

        await stream.FlushAsync();
        //stream.Close();  
    }


}
```

## Examples

## Getting support 👷
If you encounter any bugs, don't hesitate to submit an [issue](https://github.com/simplify9/DeeBee/issues). We'll get back to you promptly!

