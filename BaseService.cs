using System;
using System.Collections.Generic;
using System.IO;

namespace Office.Service
{
    public abstract class BaseService
    {
        /// <summary>
        /// 文件生成路径
        /// </summary>
        /// <param name="fileName">扩展名</param>
        /// <param name="path">路径</param>
        /// <returns></returns>
        public string GenFile(string ext, string pathFile = null)
        {
            //生成uuid
            string file =  $"{Guid.NewGuid().ToString("N")}.{ext}";
            string dir = $"{Directory.GetCurrentDirectory()}\\{DateTime.Now:yyyyMMdd}";
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            //文件路径
            pathFile = pathFile == null ? $@"{dir}\{file}" : pathFile;
            return pathFile;
        }

        #region word转换
        public abstract bool WordConvertPDF(string sourceFile, string targetFile);

        public abstract bool WordConvertExcel(string sourceFile, string targetFile);

        public abstract bool WordMerge(List<string> addLists, string targetPdf);
        #endregion

        #region excel转换
        public abstract bool ExcelConvertPDF(string sourceFile, string targetFile);

        public abstract bool ExcelConvertWord(string sourceFile, string targetFile);

        public abstract bool ExcelMerge(List<string> addLists, string targetPdf);
        #endregion

        #region powerpoint转换
        public abstract bool PowerPointConvertPDF(string sourceFile, string targetFile);

        public abstract bool PowerPointConvertWord(string sourceFile, string targetFile);

        public abstract bool PowerPointConvertExcel(string sourceFile, string targetFile);

        public abstract bool PowerPointMerge(List<string> addLists,string targetPdf);
        #endregion

        #region pdf转换
        public abstract bool PDFConvertWord(string sourceFile, string targetFile);

        public abstract bool PDFConvertExcel(string sourceFile, string targetFile);

        public abstract bool PDFMerge(List<string> addLists, string targetPdf);
        #endregion

        #region 合并
        /// <summary>
        /// 合并 
        /// </summary>
        /// <returns></returns>
        public string Merge(OfficeEnum type, List<string> adds, List<string> exists)
        {
            exists.AddRange(adds);
            string targetFile = GenFile(type.ToString());
            bool result = false;
            switch (type)
            {
                case OfficeEnum.doc:
                    result = WordMerge(exists, targetFile);
                    break;
                case OfficeEnum.xlsx:
                    result = ExcelMerge(exists, targetFile);
                    break;
                case OfficeEnum.pdf:
                    result = PDFMerge(exists, targetFile);
                    break;
            }
            try
            {
                foreach (var add in adds)
                {
                    System.IO.File.Delete(add);
                }
            }
            catch (Exception ex)
            {
                //todo
            }
            return result ? targetFile : "合并异常";
        }
        #endregion
    }
}