namespace DingExcelTool.Data
{
    using System.Collections.Generic;
    
    internal class ErrorCodeScriptFieldInfo
    {
        public string CodeStr;
        public int Code;
        public string Comment;
    }

    internal class ErrorCodeScriptInfo
    {
        public string SheetName;
        public List<ErrorCodeScriptFieldInfo> Fields;
    }
}
