namespace DingExcelTool.LanguageHandler;

using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.Concurrent;
using Data;

internal interface ILanguageHandler
{
    bool IsSkip(string inputPath);
    Task<(HashSet<string> languageHash, HashSet<string> imageHash)> GetLanguagesAsync(string inputPath);
    Task LanguageReplaceAsync(string inputPath, string outputDir, bool isClient, ScriptTypeEn scriptType, ConcurrentDictionary<string, int> languageDic, ConcurrentDictionary<string, int> imageDic, params object[] arg);
    Task LanguageRevertAsync(string inputPath, ConcurrentDictionary<int, string> languageDic, ConcurrentDictionary<int, string> imageDic);
}