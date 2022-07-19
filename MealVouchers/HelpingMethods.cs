namespace MealVouchers
{
    public class HelpingMethods
    {

        public static string FormDate(DateTime date)
        {
            var month = (date.Month + 1).ToString();
            var year = date.Year.ToString();
            return (year + ";" + month);
        }

        public static string RealMonthName(IReadOnlyList<string> names, string aimObjectName)
        {
            if (names.Contains(aimObjectName))
            {
                var index = names.IndexOf(aimObjectName);
                int newindex;
                if (index == 0)
                {
                    newindex = names.Count - 1;
                }
                else
                {
                    newindex = index - 1;
                }
                var realSheetName = names[newindex];
                return realSheetName;
            }
            else
            {
                throw new InvalidDataException("The aim sheet is missimg or has an invalid name");
            }

        }

        public static IReadOnlyList<string> FindMatchingFiles(string inputPath, string fileNameExpresion)
        {
            var result = new List<string>();
            result.AddRange(Directory.GetFiles(inputPath, fileNameExpresion, SearchOption.TopDirectoryOnly));
            foreach (string subdirectory in Directory.GetDirectories(inputPath))
            {
                result.AddRange(FindMatchingFiles(subdirectory, fileNameExpresion));
            }
            return result;
        }

    }
}
