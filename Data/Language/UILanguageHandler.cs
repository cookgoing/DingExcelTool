using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading.Tasks;

public class UILanguageHandler : OuterLanguageHandler
{
    public string Suffix => ".uxml";
    
    public bool IsSkip(string inputPath)
    {
        string extension = Path.GetExtension(inputPath);
        return extension != Suffix;
    }

    public async Task<(HashSet<string> languageHash, HashSet<string> imageHash)> GetLanguagesAsync(string inputPath, Action<string> errorLogger)
    {
        if (!File.Exists(inputPath)) throw new Exception($"[UILanguageHandler.GetLanguagesAsync] 路径不存在：{inputPath}");

        HashSet<string> languageHash = new(), imageHash = new();
        XDocument doc = XDocument.Load(inputPath);
        var nodes = doc.Descendants().Where(x =>
            (x.Name == "DingFrame.Module.TKUI.DLabel" || x.Name == "DingFrame.Module.TKUI.DButton" || x.Name == "DingFrame.Module.TKUI.DTextField")
            && (x.Attribute("Localization")?.Value != "false"));

        foreach (var node in nodes)
        {
            bool isTextField = node.Name == "DingFrame.Module.TKUI.DTextField";
            var textAttr = node.Attribute(isTextField? "Placeholder" : "text");
            if (textAttr == null || string.IsNullOrEmpty(textAttr.Value)) continue;

            string str = textAttr.Value.Replace("&#10", "\n");
            languageHash.Add(str);
        }

        return (languageHash,  imageHash);
    }

    public async Task LanguageReplaceAsync(string inputPath, ConcurrentDictionary<string, int> languageDic, ConcurrentDictionary<string, int> imageDic, Action<string> errorLogger)
    {
        if (!File.Exists(inputPath)) throw new Exception($"[UILanguageHandler.LanguageReplaceAsync] 路径不存在：{inputPath}");
        
        XDocument doc = XDocument.Load(inputPath);
        var nodes = doc.Descendants().Where(x =>
            (x.Name == "DingFrame.Module.TKUI.DLabel" || x.Name == "DingFrame.Module.TKUI.DButton" || x.Name == "DingFrame.Module.TKUI.DTextField")
            && (x.Attribute("Localization")?.Value != "false"));

        bool modified = false;
        foreach (var node in nodes)
        {
            bool isTextField = node.Name == "DingFrame.Module.TKUI.DTextField";
            var textAttr = node.Attribute(isTextField? "Placeholder" : "text");
            if (textAttr == null || string.IsNullOrEmpty(textAttr.Value)) continue;

            if (!languageDic.TryGetValue(textAttr.Value, out int hashId))
            {
                errorLogger($"文件: {inputPath}, 字段: {textAttr.Value};  没有生成对应的多语言文本");
                continue;
            }

            var textHashAttr = node.Attribute("TextHash");
            if (textHashAttr != null) textHashAttr.Value = hashId.ToString();
            else node.Add(new XAttribute("TextHash", hashId));
            
            modified = true;
        }

        if (modified) doc.Save(inputPath);
    }

    public async Task LanguageRevertAsync(string inputPath, ConcurrentDictionary<int, string> languageDic, ConcurrentDictionary<int, string> imageDic, Action<string> errorLogger)
    {
        if (!File.Exists(inputPath)) throw new Exception($"[UILanguageHandler.LanguageRevertAsync] 路径不存在：{inputPath}");
        
        XDocument doc = XDocument.Load(inputPath);
        var nodes = doc.Descendants().Where(x =>
            (x.Name == "DingFrame.Module.TKUI.DLabel" || x.Name == "DingFrame.Module.TKUI.DButton" || x.Name == "DingFrame.Module.TKUI.DTextField")
            && (x.Attribute("Localization")?.Value != "false"));

        bool modified = false;
        foreach (var node in nodes)
        {
            bool isTextField = node.Name == "DingFrame.Module.TKUI.DTextField";
            var textAttr = node.Attribute(isTextField? "Placeholder" : "text");
            if (textAttr == null || string.IsNullOrEmpty(textAttr.Value)) continue;

            var textHashAttr = node.Attribute("TextHash");
            if (textHashAttr != null)
            {
                textHashAttr.Value = "0";
                modified = true;
            }
        }
        
        if (modified) doc.Save(inputPath);
    }
}