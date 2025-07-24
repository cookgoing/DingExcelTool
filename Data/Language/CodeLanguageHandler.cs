using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading.Tasks;

public class CodeLanguageHandler : OuterLanguageHandler
{
    public string Suffix => ".cs";
    
    public bool IsSkip(string inputPath)
    {
        string extension = Path.GetExtension(inputPath);
        string fileName = Path.GetFileName(inputPath);
        return extension != Suffix || fileName == "LocalizationModule.cs";
    }

    public async Task<(HashSet<string> languageHash, HashSet<string> imageHash)> GetLanguagesAsync(string inputPath, Action<string> errorLogger)
    {
        if (!File.Exists(inputPath)) throw new Exception($"[CodeLanguageHandler.GetLanguagesAsync] 路径不存在：{inputPath}");
        
        HashSet<string> languageHash = new(), imageHash = new();
        string content = await File.ReadAllTextAsync(inputPath);
        
        var languageRegex = new Regex(@"LocalizationModule\.LanguageStr\(\s*""([^""]+)""\s*\)");
        var imageRegex = new Regex(@"LocalizationModule\.LanguageImgPath\(\s*""([^""]+)""\s*\)");
        var languageMathches = languageRegex.Matches(content);
        var imageMathches = imageRegex.Matches(content);
        foreach (Match match in languageMathches) languageHash.Add(match.Groups[1].Value);
        foreach (Match match in imageMathches) imageHash.Add(match.Groups[1].Value);
        
        return (languageHash, imageHash);
    }

    public async Task LanguageReplaceAsync(string inputPath, ConcurrentDictionary<string, int> languageDic, ConcurrentDictionary<string, int> imageDic, Action<string> errorLogger)
    {
        if (!File.Exists(inputPath)) throw new Exception($"[CodeLanguageHandler.LanguageReplaceAsync] 路径不存在：{inputPath}");
        
        string content = await File.ReadAllTextAsync(inputPath);
        bool modified = false;
        var languageRegex = new Regex(@"LocalizationModule\.LanguageStr\(\s*""([^""]+)""\s*\)");
        var imageRegex = new Regex(@"LocalizationModule\.LanguageImgPath\(\s*""([^""]+)""\s*\)");
        
        string newContent = languageRegex.Replace(content, m =>
        {
            var source = m.Groups[1].Value;
            if (!languageDic.TryGetValue(source, out int hashId))
            {
                errorLogger($"文件: {inputPath}, 字段: {source};  没有生成对应的多语言文本");
                return m.Groups[0].Value;
            }

            modified = true;
            return $"LocalizationModule.GetStr({hashId})/*{source}*/";
        });
        newContent = imageRegex.Replace(newContent, m =>
        {
            var source = m.Groups[1].Value;
            if (!imageDic.TryGetValue(source, out int hashId))
            {
                errorLogger($"文件: {inputPath}, 字段: {source};  没有生成对应的多语言文本");
                return m.Groups[0].Value;
            }
            
            modified = true;
            return $"LocalizationModule.GetImgPath({hashId})/*{source}*/";
        });

        if (modified) await File.WriteAllTextAsync(inputPath, newContent);
    }

    public async Task LanguageRevertAsync(string inputPath, ConcurrentDictionary<int, string> languageDic, ConcurrentDictionary<int, string> imageDic, Action<string> errorLogger)
    {
        if (!File.Exists(inputPath)) throw new Exception($"[CodeLanguageHandler.LanguageRevertAsync] 路径不存在：{inputPath}");
        
        string content = await File.ReadAllTextAsync(inputPath);
        bool modified = false;
        var languageRegex = new Regex(@"LocalizationModule\.GetStr\(\s*([+-]?\d+)\s*\)\s*/\*((?:.|\n)*?)\*/");
        var imageRegex = new Regex(@"LocalizationModule\.GetImgPath\(\s*([+-]?\d+)\s*\)\s*/\*((?:.|\n)*?)\*/");
        
        string newContent = languageRegex.Replace(content, m => {
            var source = m.Groups[2].Value;
            modified = true;
            return $"LocalizationModule.LanguageStr(\"{source}\")";
        });
        newContent = imageRegex.Replace(newContent, m => {
            var source = m.Groups[2].Value;
            modified = true;
            return $"LocalizationModule.LanguageImgPath(\"{source}\")";
        });
        
        if (modified) await File.WriteAllTextAsync(inputPath, newContent);
    }
}