using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace EPPlus_ManipulateExcel
{
    public class ManipulateExcel
    {

        string config_Sheetname = "Configuration";
        string filepath = @"C:\Ashok\MyProjects\EPPlus_ManipulateExcel";
        public  ManipulateExcel()
        {
        }


        

        public string ProcessDatabyTemplate(string json)
        {
            string resJson = string.Empty;
            string newFilePath = string.Empty;
            string templateName = string.Empty;

            try
            {
                Dictionary<string, object> kvp = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);

                
                templateName = kvp["TemplateToUse"].ToString();

                string templateFullPath = filepath + "\\" + templateName;

                if (!File.Exists(templateFullPath))
                {
                    Console.WriteLine("The Excel Template '" + templateName + "' is not found in the path: " + templateName);                    
                    return resJson;
                }

                string archivePath = filepath + "\\Archive";

                if (!Directory.Exists(archivePath))
                    Directory.CreateDirectory(archivePath);

                //get new filename based on config
                newFilePath = archivePath + "\\" + Path.GetFileNameWithoutExtension(templateName) + DateTime.Now.ToString("MMddyyyyHHmmss") + ".xlsx";


                //if (Path.GetExtension(templateName) == ".xls")
                //    ConvertXLSToXLSX(templateFullPath, newFilePath);
                //else
                //{
                    //save as
                    using (var app = new ExcelPackage(new FileInfo(templateFullPath)))
                    {
                        //process
                        app.SaveAs(new FileInfo(newFilePath));
                    }
                //}

                //Load Excel template
                string configJson = ExcelToJsonByEPPlus(templateFullPath);

                if (string.IsNullOrEmpty(configJson))
                {
                    Console.WriteLine("Error occured at ManipulateExcel. Config file Error");

                    return resJson;
                }

                List<TemplateConfiguration> configList = JsonConvert.DeserializeObject<List<TemplateConfiguration>>(configJson);

                //get the value from parsable
                foreach (TemplateConfiguration config in configList)
                {
                    //map the valule to config 
                    if (kvp.ContainsKey(config.ParsableName))
                        config.CellValue = kvp[config.ParsableName].ToString();
                }

                //InputDataByInterop(templateFullPath, newFilePath, configList);
                InputExcelByEPPlus(newFilePath, configList);

                resJson = RetrieveOutputFromTemplate(newFilePath, configList);

                
                Console.WriteLine("Successfully Completed Manipulating the Excel file : " + newFilePath);

            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception occured at ProcessDatabyTemplate: " + ex.Message);
            }
            finally
            {   
            }

            return resJson;
        }

        private void InputExcelByEPPlus(string newFilePath, List<TemplateConfiguration> configList)
        {
            try
            {

                //create a new Excel package from the file
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(newFilePath)))
                {
                    //get only the config for this sheet to input
                    List<TemplateConfiguration> sheetConfig = configList.FindAll(c => ((InputOrOutput)Enum.Parse(typeof(InputOrOutput), c.InputOrOutput) == InputOrOutput.Input));

                    foreach (TemplateConfiguration config in sheetConfig)
                    {
                        if ((config.CellValue == null) || (config.CellLocation == null))
                            continue;

                        if (String.IsNullOrEmpty(config.SheetName))
                        {
                            Console.WriteLine("Exception occured at InputDataByInterop: " + config.SheetName);

                            return;
                        }

                        //create an instance of the the first sheet in the loaded file
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[config.SheetName];

                        int rowIndex = int.Parse(Regex.Match(config.CellLocation, @"[0-9]+").Value);
                        int columnIndex = ColumnIndex(config.CellLocation);

                        //insert the data
                        if (config.OutputDataType == "int")
                            worksheet.Cells[rowIndex, columnIndex].Value = int.Parse(config.CellValue);
                        else if (config.OutputDataType == "float")
                            worksheet.Cells[rowIndex, columnIndex].Value = float.Parse(config.CellValue);
                        else
                            worksheet.Cells[rowIndex, columnIndex].Value = config.CellValue;

                        worksheet.Calculate();
                    }

                    //excelPackage.Workbook.Calculate();
                    //save the changes
                    excelPackage.Save();


                    excelPackage.Dispose();
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        private string ExcelToJsonByEPPlus(string templateFullPath)
        {
            string json = string.Empty;
            try
            {
                System.Data.DataTable tbl = new System.Data.DataTable();

                bool colAdded = false;

                //create a new Excel package in a memorystream
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(templateFullPath)))
                {
                    //get worksheet
                    var worksheet = excelPackage.Workbook.Worksheets[config_Sheetname];

                    if (worksheet == null)
                    {
                        Console.WriteLine("The configuration sheetname " + config_Sheetname + " is not in  " + templateFullPath);
                        return json;
                    }

                    //loop all rows
                    for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                    {
                        var row = tbl.NewRow();

                        //loop all columns in a row
                        for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                        {
                            //add the cell data to the List
                            if (worksheet.Cells[i, j].Value != null)
                            {
                                if (!colAdded)
                                    tbl.Columns.Add(worksheet.Cells[i, j].Value.ToString());

                                else
                                    row[j - worksheet.Dimension.Start.Column] = worksheet.Cells[i, j].Value.ToString();
                            }
                        }

                        if (colAdded)
                            tbl.Rows.Add(row);

                        colAdded = true;
                    }

                }

                json = JsonConvert.SerializeObject(tbl);

            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception occured at ExcelTemplateToJsonByEPPlus: " + ex.Message);
            }

            return json;
        }


        // Retrieve the value of a cell, given a file name, sheet name and address name.
        private string RetrieveOutputFromTemplate(string newFilePath, List<TemplateConfiguration> configList)
        {
            string jsonOutput = string.Empty;

            //List<OutputJson> opList = new List<OutputJson>();
            Dictionary<string, object> dic = new Dictionary<string, object>();
            try
            {
                using (var package = new ExcelPackage(new FileInfo(newFilePath)))
                {
                    package.Workbook.FullCalcOnLoad = true;

                    List<TemplateConfiguration> sheetConfig = configList.FindAll(c => (InputOrOutput)Enum.Parse(typeof(InputOrOutput), c.InputOrOutput) == InputOrOutput.Output);

                    foreach (TemplateConfiguration config in sheetConfig)
                    {
                        var theSheet = package.Workbook.Worksheets[config.SheetName];

                        if (config.CellLocation.Contains(","))
                        {
                            //to get multiple cell values for concatenated result
                            string[] strArr = config.CellLocation.Split(',');

                            theSheet.Cells[strArr[0]].Calculate();
                            config.CellValue = theSheet.Cells[strArr[0]].Value.ToString();

                            for (int i = 1; i < strArr.Length; i++)
                            {
                                theSheet.Cells[strArr[i]].Calculate();
                                config.CellValue = config.CellValue + " " + theSheet.Cells[strArr[i]].Value.ToString();
                            }
                        }
                        else
                        {
                            theSheet.Cells[config.CellLocation].Calculate();

                            config.CellValue = theSheet.Cells[config.CellLocation].Value.ToString();
                        }

                        if (string.IsNullOrEmpty(config.CellValue))
                        {
                            dic.Add(config.ParsableName, string.Empty);
                            continue;
                        }

                        if (config.OutputDataType == "float")
                        {
                            float flt = float.Parse(config.CellValue);
                            dic.Add(config.ParsableName, Math.Round(flt, 2));
                        }
                        else if (config.OutputDataType == "int")
                        {
                            int i = int.Parse(config.CellValue);
                            dic.Add(config.ParsableName, i);
                        }
                        else
                        {
                            string str = config.CellValue;
                            dic.Add(config.ParsableName, str);
                        }

                    }
                }

                jsonOutput = JsonConvert.SerializeObject(dic);

            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception occured at RetrieveOutputFromTemplate: " + ex.Message);
            }

            return jsonOutput;

        }

        private int ColumnIndex(string reference)
        {
            int ci = 0;
            reference = reference.ToUpper();
            for (int ix = 0; ix < reference.Length && reference[ix] >= 'A'; ix++)
                ci = (ci * 26) + ((int)reference[ix] - 64);
            return ci;
        }
    }
}
