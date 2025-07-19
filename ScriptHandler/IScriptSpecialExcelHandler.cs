namespace DingExcelTool.ScriptHandler
{
    using System.Threading.Tasks;
    using System.Collections.Concurrent;
    using Data;

    internal interface IScriptSpecialExcelHandler
    {
        Task GenerateErrorCodeScript(ConcurrentDictionary<string, ErrorCodeScriptInfo> errorCodeHeadDic, string frameOutputFile, string businessOutputFile);

        Task GenerateSingleScript(SingleExcelHeadInfo singleHeadInfo, string outputFile, bool isClient);
    }
}
