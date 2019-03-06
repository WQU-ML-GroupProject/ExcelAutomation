using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace WorldQuant_Module3_CSA_SkeletonCode
{
    class Program
    {
        static Excel.Workbook workbook;
        static Excel.Application app;
        
        static Excel._Worksheet oSheet;
        static Excel.Range oRng;

        static Excel.Series oSeries;
        static Excel.Range oResizeRange;
        static Excel._Chart oChart;

        static void Main(string[] args)
        {
            app = new Excel.Application();
            app.Visible = true;
            try
            {
                workbook = app.Workbooks.Open("property_pricing.xlsx", ReadOnly: false);
            }
            catch
            {
                SetUp();
            }

            var input = "";
            while (input != "x")
            {
                PrintMenu();
                input = Console.ReadLine();
                try
                {
                    var option = int.Parse(input);
                    switch (option)
                    {
                        case 1:
                            try
                            {
                                Console.Write("Enter the size: ");
                                var size = float.Parse(Console.ReadLine());
                                Console.Write("Enter the suburb: ");
                                var suburb = Console.ReadLine();
                                Console.Write("Enter the city: ");
                                var city = Console.ReadLine();
                                Console.Write("Enter the market value: ");
                                var value = float.Parse(Console.ReadLine());

                                AddPropertyToWorksheet(size, suburb, city, value);
                            }
                            catch
                            {
                                Console.WriteLine("Error: couldn't parse input");
                            }
                            break;
                        case 2:
                            Console.WriteLine("Mean price: " + CalculateMean());
                            break;
                        case 3:
                            Console.WriteLine("Price variance: " + CalculateVariance());
                            break;
                        case 4:
                            Console.WriteLine("Minimum price: " + CalculateMinimum());
                            break;
                        case 5:
                            Console.WriteLine("Maximum price: " + CalculateMaximum());
                            break;
                        default:
                            break;
                    }
                } catch { }
            }

            // save before exiting
            workbook.Save();
            workbook.Close();
            app.Quit();
        }

        static void PrintMenu()
        {
            Console.WriteLine();
            Console.WriteLine("Select an option (1, 2, 3, 4, 5) " +
                              "or enter 'x' to quit...");
            Console.WriteLine("1: Add Property");
            Console.WriteLine("2: Calculate Mean");
            Console.WriteLine("3: Calculate Variance");
            Console.WriteLine("4: Calculate Minimum");
            Console.WriteLine("5: Calculate Maximum");
            Console.WriteLine();
        }

        static void SetUp()
        {
            // TODO: Implement this method
            try
            {
                //Start Excel and get Application object.
                app = new Excel.Application();
                app.Visible = true;

                //Get a new workbook.
                workbook = (Excel.Workbook)(app.Workbooks.Add(Missing.Value));
                oSheet = (Excel.Worksheet)workbook.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "Size (in square feet)";
                oSheet.Cells[1, 2] = "Suburb";
                oSheet.Cells[1, 3] = "City";
                oSheet.Cells[1, 4] = "Market value";

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "D1").Font.Bold = true;
                oSheet.get_Range("A1", "D1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                workbook.SaveAs(@"property_pricing.xlsx");
            }
            catch (Exception e) { Console.WriteLine("Error: " + e.Message); }
        }

        static void AddPropertyToWorksheet(float size, string suburb, string city, float value)
        {
            // TODO: Implement this method
        }

        static float CalculateMean()
        {
            // TODO: Implement this method
            return 0.0f;
        }

        static float CalculateVariance()
        {
            // TODO: Implement this method
            return 0.0f;
        }

        static float CalculateMinimum()
        {
            // TODO: Implement this method
            return 0.0f;
        }

        static float CalculateMaximum()
        {
            // TODO: Implement this method
            return 0.0f;
        }
    }
}
