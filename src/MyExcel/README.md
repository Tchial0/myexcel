# MyExcel
Reading and writing Excel files with c# never got that easy.

### Installing MyExcel

You should install [MyExcel with NuGet](https://www.nuget.org/packages/MyExcel):

    Install-Package MyExcel
    
Or via the .NET Core command line interface:

    dotnet add package MyExcel

Either commands, from Package Manager Console or .NET Core CLI, will download and install MyExcel and all required dependencies.

### How to use MyExcel? Easy.

#### Writing a file
```c#
string fileLocation = @"C:\Users\Username\Desktop\MyExcelFile.xlsx";

using (ExcelWriter writer = new ExcelWriter())
{
    writer[1, 1] = "First Cell";
    writer.SaveAs(fileLocation);
}
```

#### Reading a file
```c#
using (ExcelReader reader = new ExcelReader())
{
    reader.FileLocation = fileLocation;
    Console.WriteLine(reader[1,1]);
}
```