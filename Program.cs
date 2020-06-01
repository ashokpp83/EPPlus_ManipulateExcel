using Newtonsoft.Json;
using Spire.Xls;
using System;
using System.IO;

namespace EPPlus_ManipulateExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            //ConvertXlsToXlsx
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Ashok\EWI\Maximo GRR - Study - 64875\Other Excel Templates\10538-18.xls");
            workbook.SaveToFile(@"C:\Ashok\EWI\Maximo GRR - Study - 64875\Other Excel Templates\10538-18.xlsx", ExcelVersion.Version2016);


            string inputjson = ParseJsonData();

            ManipulateExcel excelService = new ManipulateExcel();
            string json = excelService.ProcessDatabyTemplate(inputjson);

            //output data
            Console.WriteLine("ResultJson: " + JsonConvert.SerializeObject(json));
        }

        private static string ParseJsonData()
        {
            string json = string.Empty;

            try
            {
                string jsonFilePath = @"C:\Ashok\MyProjects\EPPlus_ManipulateExcel\Inputdata.json";

                json = File.ReadAllText(jsonFilePath);
            }
            catch (Exception ex)
            {

                throw ex;
            }

            return json;
        }
    }
}
