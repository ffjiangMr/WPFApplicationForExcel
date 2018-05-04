namespace WpfApplication.Worker
{
    #region using directive

    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Reflection;

    #endregion

    internal class ExcelWorker
    {
        IWorkbook workBook;
        List<Type> worksTypes = new List<Type> { typeof(HSSFWorkbook), typeof(XSSFWorkbook) };

        public IWorkbook Open(String path)
        {
            IWorkbook result = null;
            try
            {
                foreach (var type in this.worksTypes)
                {
                    if (this.OpenWorkbook(path, type))
                    {
                        result = this.workBook;
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                return result;
            }
        }

        private Boolean OpenWorkbook(String path, Type type)
        {
            Boolean result = false;
            using (var stream = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                try
                {
                    this.workBook = (IWorkbook)Activator.CreateInstance(type, stream);
                    result = true;
                }
                catch (Exception ex)
                {
                }
            }
            return result; ;
        }
    }
}
