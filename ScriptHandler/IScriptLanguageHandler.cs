namespace DingExcelTool.ScriptHandler;

using System.Threading.Tasks;

internal interface IScriptLanguageHandler
{
    public string Suffix { get; }
    Task<string> ReplaceGetterFieldInClassProperty(string source, string className, string oldReturnField, string replaceFuc);
    Task<string> ReplaceArrFieldInClassProperty(string source, string className, string arrField, string replaceFuc);
    Task<string> ReplaceMapFieldInClassProperty(string source, string className, string mapFieldName, string repalceFuc, bool replaceMapKey, bool replaceMapValue);
}