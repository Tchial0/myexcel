# MyExcel
Reading and writing Excel files with c# never got that easy.

### Let's se an example

```c#
 static void Main(string[] args)
        {
            string fileLocation = @"C:\Users\Username\Desktop\MyExcelFile.xlsx";

            using (ExcelWriter writer = new ExcelWriter())
            {
                writer[1, 1] = "First Cell";
                writer.SaveAs(fileLocation);
            }
            using(ExcelReader reader = new ExcelReader())
            {
                reader.FileLocation = fileLocation;
                Console.WriteLine(reader[1,1]);
            }
        }
```