namespace DingExcelTool.ScriptHandler;

using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using System.Linq;
using System.Threading.Tasks;
using Configure;

internal class CSharpLanguageHandler : Singleton<CSharpLanguageHandler>, IScriptLanguageHandler
{
    public string Suffix => ".cs";
    
    public async Task<string> ReplaceGetterFieldInClassProperty(string source, string className, string filed, string replaceFuc)
    {
        var root = await CSharpSyntaxTree.ParseText(source).GetRootAsync();
        string languageFieldName = $"{filed}language";
        
        var newRoot = root.ReplaceNodes(root.DescendantNodes().OfType<ClassDeclarationSyntax>()
                .Where(cls => cls.Identifier.Text == className),
            (classNode, _) =>
            {
                var stringfield = classNode.Members.OfType<FieldDeclarationSyntax>().FirstOrDefault(f => f.Declaration.Variables.First().Identifier.Text == filed);
                if (stringfield != null)
                {
                    var elementType = stringfield.Declaration.Type.ToString();
                    
                    var languageField = SyntaxFactory.FieldDeclaration(
                            SyntaxFactory.VariableDeclaration(SyntaxFactory.ParseTypeName(elementType))
                                .AddVariables(SyntaxFactory.VariableDeclarator(languageFieldName)))
                        .AddModifiers(SyntaxFactory.Token(SyntaxKind.PrivateKeyword))
                        .NormalizeWhitespace().WithTrailingTrivia(SyntaxFactory.ElasticCarriageReturnLineFeed);

                    classNode = classNode.InsertNodesAfter(stringfield, new[] { languageField });
                }
                // —— 2. 替换 get 访问器的 Body —— 
                classNode = classNode.ReplaceNodes(
                    classNode.DescendantNodes().OfType<AccessorDeclarationSyntax>()
                        .Where(acc =>
                            acc.Kind() == SyntaxKind.GetAccessorDeclaration &&
                            acc.Body != null &&
                            acc.Body.Statements.OfType<ReturnStatementSyntax>()
                                .Any(ret =>
                                    ret.Expression is IdentifierNameSyntax idn &&
                                    idn.Identifier.Text == filed)),
                    (accNode, _) =>
                    {
                        var newBody = (BlockSyntax)SyntaxFactory.ParseStatement($@"
{{
	if ({languageFieldName} != null) return {languageFieldName};

    if (int.TryParse({filed}, out int {LanguageCfg.LanguageTextImageReplaceArg})) {languageFieldName} = {replaceFuc};
    else {languageFieldName} = {filed};

    return {languageFieldName};
}}").WithTriviaFrom(accNode.Body!);
                        
                        // 用这个新块替换旧的访问器体
                        return accNode.WithBody(newBody);
                    });

                return classNode;
            });

        return newRoot.NormalizeWhitespace().ToFullString();
    }
    
    public async Task<string> ReplaceArrFieldInClassProperty(string source, string className, string arrField, string replaceFuc)
    {
        var root = await CSharpSyntaxTree.ParseText(source).GetRootAsync();
        string languageArrName = $"{arrField}language";
        
        var newRoot = root.ReplaceNodes(
            // 定位所有 class B
            root.DescendantNodes().OfType<ClassDeclarationSyntax>()
                .Where(c => c.Identifier.Text == className),
            (classNode, _) =>
            {
                // —— 1. 在 arr1_ 字段之后插入 arr1_language 字段 —— 
                var arr1Field = classNode.Members.OfType<FieldDeclarationSyntax>().FirstOrDefault(f => f.Declaration.Variables.First().Identifier.Text == arrField);
                if (arr1Field != null)
                {
                    // 保留与 arr1_ 完全相同的类型
                    var elementType = arr1Field.Declaration.Type.ToString();

                    var languageField = SyntaxFactory.FieldDeclaration(
                            SyntaxFactory.VariableDeclaration(SyntaxFactory.ParseTypeName(elementType))
                                .AddVariables(SyntaxFactory.VariableDeclarator(languageArrName)))
                        .AddModifiers(SyntaxFactory.Token(SyntaxKind.PrivateKeyword))
                        .NormalizeWhitespace().WithTrailingTrivia(SyntaxFactory.ElasticCarriageReturnLineFeed);

                    classNode = classNode.InsertNodesAfter(arr1Field, new[] { languageField });
                }
                // —— 2. 替换 get 访问器的 Body —— 
                classNode = classNode.ReplaceNodes(
                    classNode.DescendantNodes().OfType<AccessorDeclarationSyntax>()
                        .Where(acc =>
                            acc.Kind() == SyntaxKind.GetAccessorDeclaration &&
                            acc.Body != null &&
                            acc.Body.Statements.OfType<ReturnStatementSyntax>()
                                .Any(ret =>
                                    ret.Expression is IdentifierNameSyntax idn &&
                                    idn.Identifier.Text == arrField)),
                    (accNode, _) =>
                    {
                        var newBody = (BlockSyntax)SyntaxFactory.ParseStatement($@"
{{
    if ({languageArrName} != null) return {languageArrName};

    {languageArrName} = new();
    for (int i = 0; i < {arrField}.Count; ++i)
    {{
        if (int.TryParse({arrField}[i], out int {LanguageCfg.LanguageTextImageReplaceArg})) {languageArrName}.Add({replaceFuc});
        else {languageArrName}.Add({arrField}[i]);
    }}

    return {languageArrName};
}}").WithTriviaFrom(accNode.Body!);
                        
                        // 用这个新块替换旧的访问器体
                        return accNode.WithBody(newBody);
                    });

                return classNode;
            });

        return newRoot.NormalizeWhitespace().ToFullString();
    }
    
    public async Task<string> ReplaceMapFieldInClassProperty(string source, string className, string mapFieldName, string repalceFuc, bool replaceMapKey, bool replaceMapValue)
    {
        if (!replaceMapKey && !replaceMapValue) return source;
        
        var root = await CSharpSyntaxTree.ParseText(source).GetRootAsync();
        string languageMapName = $"{mapFieldName}language";

        var newRoot = root.ReplaceNodes(root.DescendantNodes().OfType<ClassDeclarationSyntax>()
                .Where(c => c.Identifier.Text == className),
            (classNode, _) =>
            {
                var mapField = classNode.Members
                    .OfType<FieldDeclarationSyntax>()
                    .FirstOrDefault(f => f.Declaration.Variables.First().Identifier.Text == mapFieldName);
                if (mapField != null)
                {
                    var mapType = mapField.Declaration.Type.ToString();

                    var languageField = SyntaxFactory.FieldDeclaration(
                            SyntaxFactory.VariableDeclaration(
                                    SyntaxFactory.ParseTypeName(mapType))
                                .AddVariables(SyntaxFactory.VariableDeclarator(languageMapName)))
                        .AddModifiers(SyntaxFactory.Token(SyntaxKind.PrivateKeyword))
                        .NormalizeWhitespace().WithTrailingTrivia(SyntaxFactory.ElasticCarriageReturnLineFeed);

                    classNode = classNode.InsertNodesAfter(mapField, new[] { languageField });
                }

                classNode = classNode.ReplaceNodes(
                    classNode.DescendantNodes().OfType<AccessorDeclarationSyntax>()
                        .Where(acc =>
                            acc.Kind() == SyntaxKind.GetAccessorDeclaration &&
                            acc.Body != null &&
                            acc.Body.Statements
                                .OfType<ReturnStatementSyntax>()
                                .Any(ret =>
                                    ret.Expression is IdentifierNameSyntax idn &&
                                    idn.Identifier.Text == mapFieldName)),
                    (accNode, _) =>
                    {
                        string key = "keyStr", value = "valueStr";
                        string keyStr = $"string {key} = " + (replaceMapKey ? $"if (int.TryParse(kv.Key, out int {LanguageCfg.LanguageTextImageReplaceArg})) {repalceFuc} else kv.Key;" : "kv.Key;");
                        string valueStr = $"string {value} = " + (replaceMapValue ? $"if (int.TryParse(kv.Value, out int {LanguageCfg.LanguageTextImageReplaceArg})) {repalceFuc} else kv.Value;" : "kv.Value;");
                        // 用 ParseBlock 方式构造新的多行 BlockSyntax
                        // 注意我们手动写上 {} 并在最外层 ParseStatement
                        var newBody = (BlockSyntax)SyntaxFactory.ParseStatement(@$"
{{
    if ({languageMapName} != null) return {languageMapName};

    {languageMapName} = new();
    foreach (var kv in {mapFieldName})
    {{
        {keyStr}
        {valueStr}
        {languageMapName}.Add({key}, {value});
    }}

    return {languageMapName};
}}").WithTriviaFrom(accNode.Body!);      // 继承原 getter 的注释/空白

                        // 用新的 body 替换旧的
                        return accNode.WithBody(newBody);
                    });

                return classNode;
            });

        return newRoot.NormalizeWhitespace().ToFullString();
    }

}