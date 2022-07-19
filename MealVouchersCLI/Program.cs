using MealVouchers;

class Program
{
    static int Main(string[] args)
    {
        if (args.Length == 1 && args[0] == "/?")
        {
            WriteHelp();
            return 0;
        }
        if (args.Length < 3 || args.Length > 4)
        {
            Console.WriteLine("Wrong number of arguments. Write \"/?\" to find out which argument needed ");
            return 1;
        }

        string inputPath = args[0];
        string fileNameExpresion = args[1];

        string aimSheet;
        if (args.Length == 3)
        {
            aimSheet = monthNames[DateTime.Now.Month - 1];
        }
        else
        {
            aimSheet = args[3];
        }

        var realSheetName = HelpingMethods.RealMonthName(monthNames, aimSheet);
        string dateCell = "A5";
        string employeeNumCell = "B1";

        string resultFilePath = args[2];
        IReadOnlyList<string> matchingFiles = HelpingMethods.FindMatchingFiles(inputPath, fileNameExpresion);

        List<ExtractedData> allEmployeeData = new();
        foreach (string filename in matchingFiles)
        {
            var excelExtractor = new ExcelExtractor(filename, realSheetName);
            if (excelExtractor.CheckCell(dateCell))
            {
                var date = excelExtractor.GetDate(dateCell);
                var realDate = HelpingMethods.FormDate(date);
                var employeeNumber = excelExtractor.GetCellValueString(employeeNumCell);
                var valueDict = excelExtractor.GetStrings(aimCells);
                var employeeData = new ExtractedData(realDate, employeeNumber, valueDict);
                allEmployeeData.Add(employeeData);
            }
        }

//        var stringWriter = new StringWriter();
        using (var writer = new StreamWriter(resultFilePath))
        {
            foreach (ExtractedData data in allEmployeeData)
            {
                data.PrintExtractedData(writer);
//                data.PrintExtractedData(stringWriter);
            }
        }
        //        Console.WriteLine(stringWriter.ToString());
        Console.WriteLine($"Successfully wrote the results for {matchingFiles.Count} files to {resultFilePath}");
        return 0;
    }

    private static void WriteHelp()
    {
        Console.WriteLine
            (@"You need to give up to 4 arguments in the following sequence:
            1.Path to a folder where the neede files are located.
            2.A name of a particular file or a pattern to find all similar files - e.g. Essensmarken_??.xlsx
            3.Path to the file with results will be placed.
            4.Month which you would like to get information about (optional: if not stated, automatically gives results for the current month).");
    }

    static readonly string[] monthNames = new string[] { "Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember" };
    static readonly string[][] aimCells = new string[][] { new string[] { "F2", "F23" }, new string[] { "G2", "G23" } };

}

