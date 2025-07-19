using System;
using System.IO;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Collections.Generic;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Platform.Storage;
using DingExcelTool.ExcelHandler;
using DingExcelTool.Utils;
using DingExcelTool.Data;
using DingExcelTool.Configure;
using DingExcelTool.LanguageHandler;

namespace DingExcelTool;

public partial class MainWindow : Window
{
    private enum ExcuteTypeEn
    { 
        ExportExcel,
        ClearOutputDir,
        RestoreDefaults,
        GenerateLanguageExcels,
        LanguageReplace,
        LanguageRevert,
    }

    private Dictionary<ScriptTypeEn, string> scriptTypeDic = new ()
    {
        {ScriptTypeEn.CSharp, "C#"},
    };

    private Dictionary<ExcuteTypeEn, string> excuteTypeDic = new()
    {
        {ExcuteTypeEn.ExportExcel, "导表"},
        {ExcuteTypeEn.ClearOutputDir, "清空缓存"},
        {ExcuteTypeEn.RestoreDefaults, "恢复默认值"},
        {ExcuteTypeEn.GenerateLanguageExcels, "生成多语言表"},
        {ExcuteTypeEn.LanguageReplace, "多语言替换+导表"},
        {ExcuteTypeEn.LanguageRevert, "多语言还原+导表"},
    };

    private bool initedUI;

    public MainWindow()
    {
        initedUI = false;
        InitializeComponent();
        RefreshImg();

        LogMessageHandler.Init(this);

        list_log.Items.Clear();
        cb_clientScriptType.Items.Clear();
        cb_serverScriptType.Items.Clear();
        cb_action.Items.Clear();
        cb_sourceLanguage.Items.Clear();
        grid_batch.IsVisible = false;

        foreach (ScriptTypeEn scriptType in Enum.GetValues<ScriptTypeEn>())
        {
            if (!scriptTypeDic.TryGetValue(scriptType, out string typeName))
            {
                LogMessageHandler.AddError($"脚本类型，没有对应的文本显示字段：{scriptType}");
                continue;
            }

            cb_clientScriptType.Items.Add(new ComboBoxItem { Content = typeName });
            cb_serverScriptType.Items.Add(new ComboBoxItem { Content = typeName });
        }

        foreach (ExcuteTypeEn excuteType in Enum.GetValues<ExcuteTypeEn>())
        {
            if (!excuteTypeDic.TryGetValue(excuteType, out string excutionName))
            {
                LogMessageHandler.AddError($"操作类型，没有对应的文本显示字段：{excuteType}");
                continue;
            }

            cb_action.Items.Add(new ComboBoxItem { Content = excutionName });
        }

        foreach (LanguageType languageType in Enum.GetValues<LanguageType>())
        {
            cb_sourceLanguage.Items.Add(new ComboBoxItem { Content = languageType });
        }

        initedUI = true;
        RefreshUI();

        cb_action.SelectedIndex = ExcelManager.Instance.Data.LastSelectedActionIdx;
        cb_sourceLanguage.SelectedIndex = (int)ExcelManager.Instance.Data.Language.SourceLanguage;
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);

        RefreshData();

        LanguageManager.Instance.Clear();
        ExcelManager.Instance.Clear();
    }

    private void RefreshImg()
    {
        var iconPath = Path.Combine(AppContext.BaseDirectory, "images", "teamLogo.png");
        var bitmap = new Bitmap(iconPath);
        double maxSize = 250;
        double width, height;
        if (bitmap.Size.Width > bitmap.Size.Height)
        {
            width = maxSize;
            height = width / bitmap.Size.AspectRatio;
        }
        else
        {
            height = maxSize;
            width = height * bitmap.Size.AspectRatio;
        }

        TeamLogoImg.Source = bitmap;
        TeamLogoImg.Width = width;
        TeamLogoImg.Height = height;
    }
    private void RefreshUI()
    {
        clientCheckBox.IsChecked = ExcelManager.Instance.Data.OutputClient;
        serverCheckBox.IsChecked = ExcelManager.Instance.Data.OutputServer;

        cb_clientScriptType.SelectedIndex = (int)ExcelManager.Instance.Data.ClientOutputInfo.ScriptType;
        cb_serverScriptType.SelectedIndex = (int)ExcelManager.Instance.Data.ServerOutputInfo.ScriptType;

        tb_excelPath.Text = ExcelManager.Instance.Data.ExcelInputRootDir;
        tb_clientPBMetaPath.Text = ExcelManager.Instance.Data.ClientOutputInfo.ProtoMetaOutputDir;
        tb_clientPBScriptPath.Text = ExcelManager.Instance.Data.ClientOutputInfo.ProtoScriptOutputDir;
        tb_clientPBDataPath.Text = ExcelManager.Instance.Data.ClientOutputInfo.ProtoDataOutputDir;
        tb_clientExcelScriptPath.Text = ExcelManager.Instance.Data.ClientOutputInfo.ExcelScriptOutputDir;
        tb_clientErrorcodeFramePath.Text = ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeFrameDir;
        tb_clientErrorcodeBusinessPath.Text = ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeBusinessDir;

        tb_serverPBMetaPath.Text = ExcelManager.Instance.Data.ServerOutputInfo.ProtoMetaOutputDir;
        tb_serverPBScriptPath.Text = ExcelManager.Instance.Data.ServerOutputInfo.ProtoScriptOutputDir;
        tb_serverPBDataPath.Text = ExcelManager.Instance.Data.ServerOutputInfo.ProtoDataOutputDir;
        tb_serverExcelScriptPath.Text = ExcelManager.Instance.Data.ServerOutputInfo.ExcelScriptOutputDir;
        tb_serverErrorcodeFramePath.Text = ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeFrameDir;
        tb_serverErrorcodeBusinessPath.Text = ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeBusinessDir;
        
        tb_languageExcelDir.Text = ExcelManager.Instance.Data.Language.LanguageExcelDir;
        tb_languageScriptRootDir.Text = ExcelManager.Instance.Data.Language.ScriptRootDir;
        tb_languageUIRootDir.Text = ExcelManager.Instance.Data.Language.UIRootDir;
        tb_languageBackupDir.Text = ExcelManager.Instance.Data.Language.LanguageBackupDir;
        cb_sourceLanguage.SelectedIndex = (int)ExcelManager.Instance.Data.Language.SourceLanguage;
        tb_languageReplaceMethod.Text = ExcelManager.Instance.Data.Language.LanguageReplaceMethod;
        tb_imageReplaceMethod.Text = ExcelManager.Instance.Data.Language.ImageReplaceMethod;
    }
    private void RefreshData()
    {
        ExcelManager.Instance.Data.OutputClient = clientCheckBox.IsChecked ?? false;
        ExcelManager.Instance.Data.OutputServer = serverCheckBox.IsChecked ?? false;

        ExcelManager.Instance.Data.ClientOutputInfo.ScriptType = (ScriptTypeEn)cb_clientScriptType.SelectedIndex;
        ExcelManager.Instance.Data.ServerOutputInfo.ScriptType = (ScriptTypeEn)cb_serverScriptType.SelectedIndex;

        ExcelManager.Instance.Data.ExcelInputRootDir = tb_excelPath.Text;
        ExcelManager.Instance.Data.ClientOutputInfo.ProtoMetaOutputDir = tb_clientPBMetaPath.Text;
        ExcelManager.Instance.Data.ClientOutputInfo.ProtoScriptOutputDir = tb_clientPBScriptPath.Text;
        ExcelManager.Instance.Data.ClientOutputInfo.ProtoDataOutputDir = tb_clientPBDataPath.Text;
        ExcelManager.Instance.Data.ClientOutputInfo.ExcelScriptOutputDir = tb_clientExcelScriptPath.Text;
        ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeFrameDir = tb_clientErrorcodeFramePath.Text;
        ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeBusinessDir = tb_clientErrorcodeBusinessPath.Text;

        ExcelManager.Instance.Data.ServerOutputInfo.ProtoMetaOutputDir = tb_serverPBMetaPath.Text;
        ExcelManager.Instance.Data.ServerOutputInfo.ProtoScriptOutputDir = tb_serverPBScriptPath.Text;
        ExcelManager.Instance.Data.ServerOutputInfo.ProtoDataOutputDir = tb_serverPBDataPath.Text;
        ExcelManager.Instance.Data.ServerOutputInfo.ExcelScriptOutputDir = tb_serverExcelScriptPath.Text;
        ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeFrameDir = tb_serverErrorcodeFramePath.Text;
        ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeBusinessDir = tb_serverErrorcodeBusinessPath.Text;
        
        ExcelManager.Instance.Data.Language.LanguageExcelDir = tb_languageExcelDir.Text;
        ExcelManager.Instance.Data.Language.ScriptRootDir = tb_languageScriptRootDir.Text;
        ExcelManager.Instance.Data.Language.UIRootDir = tb_languageUIRootDir.Text;
        ExcelManager.Instance.Data.Language.LanguageBackupDir = tb_languageBackupDir.Text;
        ExcelManager.Instance.Data.Language.SourceLanguage = (LanguageType)cb_sourceLanguage.SelectedIndex;
        ExcelManager.Instance.Data.Language.LanguageReplaceMethod = tb_languageReplaceMethod.Text;
        ExcelManager.Instance.Data.Language.ImageReplaceMethod = tb_imageReplaceMethod.Text;
    }

    private async void SelectorAction(TextBox pathTextor, string desc)
    {
        var folderDialog = new FolderPickerOpenOptions
        {
            Title = desc,
            AllowMultiple = false
        };

        var list = await StorageProvider.OpenFolderPickerAsync(folderDialog);
        if (list.Count > 0) pathTextor.Text = list[0].Path.LocalPath;
    }
    public int AddLogItem(string logStr, LogType logType)
    {
        return list_log.Items.Add(new ListBoxItem
        {
            Content = logStr,
            Foreground = logType == LogType.Error ? Brushes.Red : logType == LogType.Warn ? Brushes.Orange : Brushes.Black,
        });
    }
    public void MoveLogListScroll(int idx = -1)
    {
        if (idx == -1)
        {
            MoveLogListScroll(list_log.Items.Count - 1);
            return;
        }

        if (idx < 0 || idx >= list_log.Items.Count) return;

        list_log.ScrollIntoView(idx);
    }

    private void CheckBox_clientStateChanged(object sender, RoutedEventArgs e) { if (initedUI) ExcelManager.Instance.Data.OutputClient = clientCheckBox.IsChecked ?? false; }
    private void CheckBox_serverStateChanged(object sender, RoutedEventArgs e) { if (initedUI) ExcelManager.Instance.Data.OutputServer = serverCheckBox.IsChecked ?? false; }

    private void Btn_excelFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_excelPath, "选择表格文件夹"); ExcelManager.Instance.Data.ExcelInputRootDir = tb_excelPath.Text; }
    private void Btn_clientPBmetaFolderSelector(object sender, RoutedEventArgs e) {  SelectorAction(tb_clientPBMetaPath, "选择客户端的proto meta的导出文件夹");  ExcelManager.Instance.Data.ClientOutputInfo.ProtoMetaOutputDir = tb_clientPBMetaPath.Text; }
    private void Btn_clientPBScriptFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_clientPBScriptPath, "选择客户端的proto script的导出文件夹"); ExcelManager.Instance.Data.ClientOutputInfo.ProtoScriptOutputDir = tb_clientPBScriptPath.Text; }
    private void Btn_clientPBDataFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_clientPBDataPath, "选择客户端的proto data的导出文件夹"); ExcelManager.Instance.Data.ClientOutputInfo.ProtoDataOutputDir = tb_clientPBDataPath.Text; }
    private void Btn_clientExcelScriptFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_clientExcelScriptPath, "选择客户端的excel script的导出文件夹"); ExcelManager.Instance.Data.ClientOutputInfo.ExcelScriptOutputDir = tb_clientExcelScriptPath.Text; }
    private void Btn_clientECFrameFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_clientErrorcodeFramePath, "选择客户端的error code frame的导出文件夹"); ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeFrameDir = tb_clientErrorcodeFramePath.Text; }
    private void Btn_clientECBusinessFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_clientErrorcodeBusinessPath, "选择客户端的error code business的导出文件夹"); ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeBusinessDir = tb_clientErrorcodeBusinessPath.Text; }
    private void Btn_serverPBMetaFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_serverPBMetaPath, "选择服务器的proto meta的导出文件夹"); ExcelManager.Instance.Data.ServerOutputInfo.ProtoMetaOutputDir = tb_serverPBMetaPath.Text; }
    private void Btn_serverPBScriptFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_serverPBScriptPath, "选择服务器的proto script的导出文件夹"); ExcelManager.Instance.Data.ServerOutputInfo.ProtoScriptOutputDir = tb_serverPBScriptPath.Text; }
    private void Btn_serverPBDataFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_serverPBDataPath, "选择服务器的proto data的导出文件夹"); ExcelManager.Instance.Data.ServerOutputInfo.ProtoDataOutputDir = tb_serverPBDataPath.Text; }
    private void Btn_serverExcelScriptFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_serverExcelScriptPath, "选择服务器的excel script的导出文件夹"); ExcelManager.Instance.Data.ServerOutputInfo.ExcelScriptOutputDir = tb_serverExcelScriptPath.Text; }
    private void Btn_serverECFrameFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_serverErrorcodeFramePath, "选择服务器的error code frame的导出文件夹"); ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeFrameDir = tb_serverErrorcodeFramePath.Text; }
    private void Btn_serverECBusinessFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_serverErrorcodeBusinessPath, "选择服务器的error code business的导出文件夹"); ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeBusinessDir = tb_serverErrorcodeBusinessPath.Text; }
    private void Btn_preProcessSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_preProcessPath, "选择前处理的文件"); ExcelManager.Instance.Data.PreHanleProgramFile = tb_preProcessPath.Text; }
    private void Btn_aftProcessSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_aftProcessPath, "选择后处理的文件"); ExcelManager.Instance.Data.AftHanleProgramFile = tb_aftProcessPath.Text; }
    private void CB_clientScriptTypeChanged(object sender, SelectionChangedEventArgs e) { if (initedUI) ExcelManager.Instance.Data.ClientOutputInfo.ScriptType = (ScriptTypeEn)cb_clientScriptType.SelectedIndex; }
    private void CB_serverScriptTypeChanged(object sender, SelectionChangedEventArgs e) { if (initedUI) ExcelManager.Instance.Data.ServerOutputInfo.ScriptType = (ScriptTypeEn)cb_serverScriptType.SelectedIndex; }
    private void CB_actionTypeChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!initedUI) return;

        ExcelManager.Instance.Data.LastSelectedActionIdx = cb_action.SelectedIndex; 
    }

    private void Btn_languageExcelFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_excelPath, "选择多语言表文件夹"); ExcelManager.Instance.Data.Language.LanguageExcelDir = tb_languageExcelDir.Text; }
    private void Btn_languageScriptFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_excelPath, "选择多语言代码文件夹"); ExcelManager.Instance.Data.Language.ScriptRootDir = tb_languageScriptRootDir.Text; }
    private void Btn_languageUIFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_excelPath, "选择多语言UI文件夹"); ExcelManager.Instance.Data.Language.UIRootDir = tb_languageUIRootDir.Text; }
    private void Btn_languageBackupFolderSelector(object sender, RoutedEventArgs e) { SelectorAction(tb_excelPath, "选择多语言备份文件夹"); ExcelManager.Instance.Data.Language.LanguageBackupDir = tb_languageBackupDir.Text; }
    private void CB_sourceLanguageChanged(object sender, SelectionChangedEventArgs e) { if (initedUI) ExcelManager.Instance.Data.Language.SourceLanguage = (LanguageType)cb_sourceLanguage.SelectedIndex; }
    private void TB_languageReplaceMethodChanged(object? sender, TextChangedEventArgs e) => ExcelManager.Instance.Data.Language.LanguageReplaceMethod = tb_languageReplaceMethod.Text;
    private void TB_imageReplaceMethodChanged(object? sender, TextChangedEventArgs e) => ExcelManager.Instance.Data.Language.ImageReplaceMethod = tb_imageReplaceMethod.Text;
    
    private async void Btn_excute(object sender, RoutedEventArgs eventArg)
    {
        try
        {
            ExcuteTypeEn excutionType = (ExcuteTypeEn)ExcelManager.Instance.Data.LastSelectedActionIdx;
            switch (excutionType)
            {
                case ExcuteTypeEn.ExportExcel:
                    if (!string.IsNullOrEmpty(tb_preProcessPath.Text)) ExcelUtil.ExcuteProgramFile(tb_preProcessPath.Text, tb_preProcessArgs.Text);

                    Stopwatch stopwatch = Stopwatch.StartNew();
                    RefreshData();
                    bool result = await ExportExcel();
                    stopwatch.Stop();

                    if (!string.IsNullOrEmpty(tb_aftProcessPath.Text)) ExcelUtil.ExcuteProgramFile(tb_preProcessPath.Text, tb_aftProcessArgs.Text);

                    if (result) LogMessageHandler.AddInfo($"导表 完成！花费时间：{stopwatch.ElapsedMilliseconds / 1000f}s");
                    else LogMessageHandler.AddError("导表 出现错误，请查看具体问题");
                    break;
                case ExcuteTypeEn.ClearOutputDir:
                    await ClearOutputDir();
                    LogMessageHandler.AddInfo("清理缓存 完成");
                    break;
                case ExcuteTypeEn.RestoreDefaults:
                    await RestoreDefault();
                    LogMessageHandler.AddInfo("恢复默认值 完成");
                    break;
                case ExcuteTypeEn.GenerateLanguageExcels:
                    stopwatch = Stopwatch.StartNew();
                    RefreshData();
                    result = await LanguageManager.Instance.ExcelGenerate();
                    if (result) LogMessageHandler.AddInfo($"生成多语言表 完成!花费时间：{stopwatch.ElapsedMilliseconds / 1000f}s");
                    else LogMessageHandler.AddError("生成多语言表 出现错误，请查看具体问题");
                    break;
                case ExcuteTypeEn.LanguageReplace:
                    if (!ExcelManager.Instance.Data.OutputClient && !ExcelManager.Instance.Data.OutputServer)
                    {
                        LogMessageHandler.AddWarn("没有输出路径，不进行相关的导出操作");
                        return;
                    }
                    await ClearOutputDir();
                    
                    stopwatch = Stopwatch.StartNew();
                    RefreshData();
                    result = await LanguageManager.Instance.LanguageReplaceAndExportExcel();
                    if (result) LogMessageHandler.AddInfo($"替换多语言表 完成!花费时间：{stopwatch.ElapsedMilliseconds / 1000f}s");
                    else LogMessageHandler.AddError("替换多语言表 出现错误，请查看具体问题");
                    break;
                case ExcuteTypeEn.LanguageRevert:
                    if (!ExcelManager.Instance.Data.OutputClient && !ExcelManager.Instance.Data.OutputServer)
                    {
                        LogMessageHandler.AddWarn("没有输出路径，不进行相关的导出操作");
                        return;
                    }
                    await ClearOutputDir();
                    
                    stopwatch = Stopwatch.StartNew();
                    RefreshData();
                    result = await LanguageManager.Instance.LanguageRevertAndExportExcel();
                    if (result) LogMessageHandler.AddInfo($"还原多语言表 完成!花费时间：{stopwatch.ElapsedMilliseconds / 1000f}s");
                    else LogMessageHandler.AddError("还原多语言表 出现错误，请查看具体问题");
                    break;
            }
        }
        catch (Exception e)
        {
            LogMessageHandler.LogException(e);
        }
        finally
        {
            LanguageManager.Instance.Reset();
            ExcelManager.Instance.Reset();
        }
    }
    
    private async Task<bool> ExportExcel()
    {
        if (!ExcelManager.Instance.Data.OutputClient && !ExcelManager.Instance.Data.OutputServer)
        {
            LogMessageHandler.AddWarn($"没有输出路径，不进行相关的导出操作");
            return false;
        }

        await ClearOutputDir();

        bool result = await ExcelManager.Instance.GenerateExcelHeadInfo();
        result &= await ExcelManager.Instance.GenerateProtoMeta();
        result &= await ExcelManager.Instance.GenerateProtoScript();
        result &= await ExcelManager.Instance.GenerateProtoData();
        result &= await ExcelManager.Instance.GenerateExcelScript();

        return result;
    }
    private async Task ClearOutputDir()
    {
        if (ExcelManager.Instance.Data.OutputClient)
        {
            ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ClientOutputInfo.ProtoMetaOutputDir);
            ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ClientOutputInfo.ProtoScriptOutputDir);
            ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ClientOutputInfo.ProtoDataOutputDir);
            ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ClientOutputInfo.ExcelScriptOutputDir);

            string frameOutputDir = ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeFrameDir;
            string businessOutputDir = ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeBusinessDir;
            string fameFilePath = Path.Combine(frameOutputDir, SpecialExcelCfg.ErrorCodeFrameScriptFileName);
            string businessFilePath = Path.Combine(businessOutputDir, SpecialExcelCfg.ErrorCodeFrameScriptFileName);

            if (Path.Exists(fameFilePath)) File.Delete(fameFilePath);
            if (Path.Exists(businessFilePath)) File.Delete(businessFilePath);
        }

        if (ExcelManager.Instance.Data.OutputServer)
        {
            ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ServerOutputInfo.ProtoMetaOutputDir);
            ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ServerOutputInfo.ProtoScriptOutputDir);
            ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ServerOutputInfo.ProtoDataOutputDir);
            ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ServerOutputInfo.ExcelScriptOutputDir);

            string frameOutputDir = ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeFrameDir;
            string businessOutputDir = ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeBusinessDir;
            string fameFilePath = Path.Combine(frameOutputDir, SpecialExcelCfg.ErrorCodeFrameScriptFileName);
            string businessFilePath = Path.Combine(businessOutputDir, SpecialExcelCfg.ErrorCodeFrameScriptFileName);

            if (Path.Exists(fameFilePath)) File.Delete(fameFilePath);
            if (Path.Exists(businessFilePath)) File.Delete(businessFilePath);
        }

        await Task.CompletedTask;
    }
    private async Task RestoreDefault()
    {
        ExcelManager.Instance.ResetDefaultData();

        RefreshUI();

        await Task.CompletedTask;
    }

}