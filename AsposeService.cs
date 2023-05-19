using System;
using System.Collections.Generic;
using System.Text;

namespace Office.Service
{
    /// <summary>
    ///  Aspose服务
    /// </summary>
    public class AsposeService : BaseService
    {
        private readonly OfficeSettings settings;
        public AsposeService(OfficeSettings options) => this.settings = options;

        public override bool ExcelConvertPDF(string sourceFile, string targetFile)
        {
            throw new NotImplementedException();
        }

        public override bool ExcelConvertWord(string sourceFile, string targetFile)
        {
            throw new NotImplementedException();
        }

        public override bool ExcelMerge(List<string> addLists, string targetPdf)
        {
            throw new NotImplementedException();
        }

        public override bool PDFConvertExcel(string sourceFile, string targetFile)
        {
            throw new NotImplementedException();
        }

        public override bool PDFConvertWord(string sourceFile, string targetFile)
        {
            throw new NotImplementedException();
        }

        public override bool PDFMerge(List<string> addLists, string targetPdf)
        {
            throw new NotImplementedException();
        }

        public override bool PowerPointConvertExcel(string sourceFile, string targetFile)
        {
            throw new NotImplementedException();
        }

        public override bool PowerPointConvertPDF(string sourceFile, string targetFile)
        {
            throw new NotImplementedException();
        }

        public override bool PowerPointConvertWord(string sourceFile, string targetFile)
        {
            throw new NotImplementedException();
        }

        public override bool PowerPointMerge(List<string> addLists, string targetPdf)
        {
            throw new NotImplementedException();
        }

        public override bool WordConvertExcel(string sourceFile, string targetFile)
        {
            throw new NotImplementedException();
        }

        public override bool WordConvertPDF(string sourceFile, string targetFile)
        {
            throw new NotImplementedException();
        }

        public override bool WordMerge(List<string> addLists,  string targetPdf)
        {
            throw new NotImplementedException();
        }
    }
}
