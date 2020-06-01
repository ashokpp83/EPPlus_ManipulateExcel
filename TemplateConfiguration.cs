using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlus_ManipulateExcel
{

    internal enum InputOrOutput
    {
        Input,
        Output
    }

    internal class TemplateConfiguration
    {

        public string ParsableName { get; set; }

        public string SheetName { get; set; }

        public string InputOrOutput { get; set; }

        public string OutputDataType { get; set; }

        public string CellLocation { get; set; }

        public string CellValue { get; set; }

    }
}
