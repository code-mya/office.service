using System;
using System.Collections.Generic;
using System.Text;

namespace Office.Service
{
    public class OfficeSettings
    {
        public TypeEnum Type { get; set; }
    }

    public enum TypeEnum
    {
        Manual,
        Spir,
        Aspose
    }

    public enum OfficeEnum
    {
        xlsx,
        doc,
        ppt,
        pdf
    }
}