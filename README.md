#NOTE:
This package was copied from https://github.com/eraydin/EPPlus.Core.Extensions
The owner no longer supports it with no reason given. The codebase worked with EPPlus 6.x, so we changed the namespaces and package name

<img src="https://user-images.githubusercontent.com/1365728/127748628-47575d74-a2fb-4539-a31e-74d8b435fc21.png" width="30%" >

[Nuget](https://img.shields.io/nuget/v/OLT.Extensions.EPPlus)](https://www.nuget.org/packages/OLT.Extensions.EPPlus)

##### Basic examples:

```cs
public class PersonDto
    {
        [ExcelTableColumn("First name")]
        [Required(ErrorMessage = "First name cannot be empty.")]
        [MaxLength(50, ErrorMessage = "First name cannot be more than {1} characters.")]
        public string FirstName { get; set; }

        [ExcelTableColumn(columnName = "Last name", isOptional = true)]
        public string LastName { get; set; }

        [ExcelTableColumn(3)]
        [Range(1900, 2050, ErrorMessage = "Please enter a value bigger than {1}")]
        public int YearBorn { get; set; }

        public decimal NotMapped { get; set; }

        [ExcelTableColumn(isOptional = true)]
        public decimal OptionalColumn1 { get; set; }

        [ExcelTableColumn(columnIndex=999, isOptional = true)]
        public decimal OptionalColumn2 { get; set; }
    }
```

- Converting from Excel to list of objects

```cs
    // Direct usage:
        excelPackage.ToList<PersonDto>(configuration => configuration.SkipCastingErrors());

    // Specific worksheet:
        excelPackage.GetWorksheet("Persons").ToList<PersonDto>();
```

- From a list of objects to Excel package

```cs
    List<PersonDto> persons = new List<PersonDto>();

    // Convert list into ExcelPackage
        ExcelPackage excelPackage = persons.ToExcelPackage();

    // Convert list into byte array
        byte[] excelPackageXlsx = persons.ToXlsx();


    // Generate ExcelPackage with configuration

    List<PersonDto> pre50 = persons.Where(x => x.YearBorn < 1950).ToList();
    List<PersonDto> post50 = persons.Where(x => x.YearBorn >= 1950).ToList();

    ExcelPackage excelPackage = pre50.ToWorksheet("< 1950")
                             .WithConfiguration(configuration => configuration.WithColumnConfiguration(x => x.AutoFit()))
                             .WithColumn(x => x.FirstName, "First Name")
                             .WithColumn(x => x.LastName, "Last Name")
                             .WithColumn(x => x.YearBorn, "Year of Birth")
                             .WithTitle("< 1950")
                             .NextWorksheet(post50, "> 1950")
                             .WithColumn(x => x.LastName, "Last Name")
                             .WithColumn(x => x.YearBorn, "Year of Birth")
                             .WithTitle("> 1950")
                             .ToExcelPackage();
```

- Generating an Excel template from ExcelWorksheetAttribute marked classes

```cs
    [ExcelWorksheet("Stocks")]
    public class StocksDto
    {
        [ExcelTableColumn("SKU")]
        public string Barcode { get; set; }

        [ExcelTableColumn]
        public int Quantity { get; set; }
    }

    // To ExcelPackage
    ExcelPackage excelPackage = Assembly.GetExecutingAssembly().GenerateExcelPackage(nameof(StocksDto));

    // To ExcelWorksheet
    using(var excelPackage = new ExcelPackage()){

        ExcelWorksheet worksheet = excelPackage.GenerateWorksheet(Assembly.GetExecutingAssembly(), nameof(StocksDto));

    }
```
