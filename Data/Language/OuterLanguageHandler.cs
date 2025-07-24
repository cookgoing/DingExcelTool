using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Collections.Concurrent;

public interface OuterLanguageHandler
{
    string Suffix { get; }
    bool IsSkip(string inputPath);
    Task<(HashSet<string> languageHash, HashSet<string> imageHash)> GetLanguagesAsync(string inputPath, Action<string> errorLogger);
    Task LanguageReplaceAsync(string inputPath, ConcurrentDictionary<string, int> languageDic, ConcurrentDictionary<string, int> imageDic, Action<string> errorLogger);
    Task LanguageRevertAsync(string inputPath, ConcurrentDictionary<int, string> languageDic, ConcurrentDictionary<int, string> imageDic, Action<string> errorLogger);
}