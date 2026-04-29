using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;

            // Converts inches to points
            double inches = 4.5;
            double points = application.InchesToPoints(inches);

            Console.WriteLine($"{inches} inches is equal to {points} points.");
        }
    }
}