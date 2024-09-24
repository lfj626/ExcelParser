using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Reflection;
using System.CodeDom;
using System.CodeDom.Compiler;

namespace MyTurn.ExcelParser
{
    public class ScriptMaker
    {
        public static void MakeScript(string FilePath, DataSet dataSet)
        {
            // 타겟 유닛 생성 
            CodeCompileUnit targetUnit = new CodeCompileUnit();

            // 네임스페이스
            CodeNamespace samples = new CodeNamespace();
            samples.Imports.Add(new CodeNamespaceImport("System"));
            samples.Imports.Add(new CodeNamespaceImport("System.Collections.Generic"));
            samples.Imports.Add(new CodeNamespaceImport("System.Linq"));
            samples.Imports.Add(new CodeNamespaceImport("System.Reflection"));
            samples.Imports.Add(new CodeNamespaceImport("Newtonsoft.Json"));
            samples.Imports.Add(new CodeNamespaceImport("Newtonsoft.Json.Linq"));
            samples.Imports.Add(new CodeNamespaceImport("Newtonsoft.Json.Serialization"));

            // 형식 선언
            CodeTypeDeclaration DataEnum = new CodeTypeDeclaration("TableType");

            // Enum
            DataEnum.IsEnum = true;
            DataEnum.TypeAttributes = TypeAttributes.Public;
            for (int i = 0; i < dataSet.Tables.Count; i++)
            {
                // 필드 선언
                CodeMemberField enumMember = new CodeMemberField("TableType", dataSet.Tables[i].TableName);
                // 필드 초기화
                enumMember.InitExpression = new CodePrimitiveExpression(i);
                // Enum에 필드 추가
                DataEnum.Members.Add(enumMember);
            }
            DataEnum.Members.Add(new CodeMemberField("TableType", "Max"));

            // Enum 추가
            samples.Types.Add(DataEnum);
            // 네임스페이스 추가
            targetUnit.Namespaces.Add(samples);

            // Json Class
            CodeTypeDeclaration JsonClass = new CodeTypeDeclaration("PrivateResolver : DefaultContractResolver");
            JsonClass.IsClass = true;
            JsonClass.TypeAttributes = TypeAttributes.Public;

            CodeMemberMethod settingMethod = new CodeMemberMethod();
            settingMethod.Attributes = MemberAttributes.Family | MemberAttributes.Override;
            settingMethod.Name = "CreateProperty";

            settingMethod.Parameters.Add(new CodeParameterDeclarationExpression("MemberInfo", "member"));
            settingMethod.Parameters.Add(new CodeParameterDeclarationExpression("MemberSerialization", "memberSerialization"));
            settingMethod.ReturnType = new CodeTypeReference("JsonProperty");
            settingMethod.Statements.Add(new CodeVariableDeclarationStatement("JsonProperty", "prop", new CodeVariableReferenceExpression("base.CreateProperty(member, memberSerialization)")));

            var ifContext = new CodeConditionStatement(
                new CodeVariableReferenceExpression("!prop.Writable"),
                new CodeVariableDeclarationStatement("var", "property", new CodeVariableReferenceExpression("member as PropertyInfo")),
                new CodeVariableDeclarationStatement("var", "hasPrivateSetter", new CodeVariableReferenceExpression("property?.GetSetMethod(true) != null")),
                new CodeAssignStatement(
                    new CodeFieldReferenceExpression(null, "prop.Writable"), new CodeFieldReferenceExpression(null, "hasPrivateSetter")));

            settingMethod.Statements.Add(ifContext);
            settingMethod.Statements.Add(new CodeMethodReturnStatement(new CodeArgumentReferenceExpression("prop")));

            JsonClass.Members.Add(settingMethod);
            samples.Types.Add(JsonClass);

            // 형식 선언
            CodeTypeDeclaration GameDataClass = new CodeTypeDeclaration("GameTable");
            // Class
            GameDataClass.IsClass = true;
            GameDataClass.TypeAttributes = TypeAttributes.Public;

            // 필드 선언
            // int 형식의 "DataVersion" 필드 선언
            CodeMemberField versionProperty = new CodeMemberField(new CodeTypeReference(typeof(int)), "DataVersion");
            versionProperty.Attributes = MemberAttributes.Private | MemberAttributes.Final | MemberAttributes.Static;
            GameDataClass.Members.Add(versionProperty);

            // 프로퍼티
            CodeMemberProperty widthProperty = new CodeMemberProperty();
            widthProperty.Attributes = MemberAttributes.Public | MemberAttributes.Final | MemberAttributes.Static;
            widthProperty.Name = "Version";
            widthProperty.HasGet = true;
            widthProperty.Type = new CodeTypeReference(typeof(int));
            widthProperty.GetStatements.Add(new CodeMethodReturnStatement(
                new CodeFieldReferenceExpression(
                    null, "DataVersion")));
            GameDataClass.Members.Add(widthProperty);

            // 메소드 
            CodeMemberMethod toLoadMethod = new CodeMemberMethod();
            toLoadMethod.Attributes = MemberAttributes.Public | MemberAttributes.Static | MemberAttributes.Final;
            toLoadMethod.Name = "ParsedData";

            // 파라미터
            toLoadMethod.Parameters.Add(new CodeParameterDeclarationExpression("JArray", "ListDatas"));
            toLoadMethod.Parameters.Add(new CodeParameterDeclarationExpression(typeof(int), "Version"));

            // 리턴 타입
            toLoadMethod.ReturnType = new CodeTypeReference("Dictionary<TableType, object>");

            // Dictionary<TableType, object> tempData = new Dictinary<TableType, object>();
            CodeObjectCreateExpression objectCreate =
             new CodeObjectCreateExpression(
             new CodeTypeReference("Dictionary<TableType, object>"));

            toLoadMethod.Statements.Add(new CodeVariableDeclarationStatement(
                new CodeTypeReference("Dictionary<TableType, object>"), "tempData",
                objectCreate));

            CodeObjectCreateExpression jsonCreate = new CodeObjectCreateExpression(new CodeTypeReference("JsonSerializerSettings"));
            toLoadMethod.Statements.Add(new CodeVariableDeclarationStatement(new CodeTypeReference("JsonSerializerSettings"), "settings", jsonCreate));

            CodeObjectCreateExpression resolverCreate = new CodeObjectCreateExpression(new CodeTypeReference("PrivateResolver"));
            toLoadMethod.Statements.Add(new CodeAssignStatement(new CodeFieldReferenceExpression(null, "settings.ContractResolver"), resolverCreate));

            // 조건
            string conditionFormat = @"CurTableName == ""{0}""";
            List<CodeStatement> loopState = new List<CodeStatement>();

            // string CurTableName = ListDatas[count]["Key"].Value<string>();
            loopState.Add(new CodeVariableDeclarationStatement(typeof(string), "CurTableName", new CodeVariableReferenceExpression("ListDatas[count][\"Key\"].Value<string>()")));
            // string CurTableData = ListDatas[count]["Value"].Value<string>();
            loopState.Add(new CodeVariableDeclarationStatement(typeof(string), "CurTableData", new CodeVariableReferenceExpression("ListDatas[count][\"Value\"].Value<string>()")));

            for (int count = 0; count < dataSet.Tables.Count; ++count)
            {
                loopState.Add(new CodeConditionStatement(
                    // if(CurTableName == "tablename")
                    new CodeVariableReferenceExpression(string.Format(conditionFormat, dataSet.Tables[count].TableName)),
                    new CodeStatement[]
                    {
                        // tempData.Add(Enum.Parse(typeof(GameTable), CurTableName, false), tablename.LoadData(ListDatas[count])
                        new CodeExpressionStatement(
                            new CodeMethodInvokeExpression(
                                new CodeFieldReferenceExpression(null, "tempData"), "Add",new CodeExpression[]
                                {
                                    // Enum.Parse(typeof(GameData), tablename, false)
                                    new CodeMethodInvokeExpression(
                                        new CodeFieldReferenceExpression(null, "(TableType)Enum"), "Parse",
                                        new CodeExpression[]
                                        {
                                            new CodeVariableReferenceExpression("typeof(TableType)"),
                                            new CodeVariableReferenceExpression("CurTableName"),
                                            new CodePrimitiveExpression(false),
                                        }),

                                    // tablename.LoadData(ListDatas[count])
                                    new CodeMethodInvokeExpression(
                                        new CodeFieldReferenceExpression(null, $"JsonConvert"), $"DeserializeObject<Dictionary<int,{dataSet.Tables[count].TableName}>>",
                                        new CodeVariableReferenceExpression("CurTableData, settings"))
                                }))
                    }));
            }

            // for( int count = 0; count < ListDatas.count; ++count)
            CodeIterationStatement forLoop = new CodeIterationStatement(
                // int count = 1;
                new CodeAssignStatement(new CodeVariableReferenceExpression("int count"), new CodePrimitiveExpression(0)),
                // count < ListDatas.Count;
                new CodeBinaryOperatorExpression(new CodeVariableReferenceExpression("count"),
                CodeBinaryOperatorType.LessThan, new CodeVariableReferenceExpression("ListDatas.Count")),
                // ++count;
                new CodeExpressionStatement(new CodeVariableReferenceExpression("++count")),
                // loopState
                loopState.ToArray());

            toLoadMethod.Statements.Add(forLoop);

            toLoadMethod.Statements.Add(new CodeAssignStatement(
                new CodeFieldReferenceExpression(null, "DataVersion"), new CodeFieldReferenceExpression(null, "Version")));

            // return tempData;
            toLoadMethod.Statements.Add(new CodeMethodReturnStatement(new CodeArgumentReferenceExpression("tempData")));
            GameDataClass.Members.Add(toLoadMethod);

            // 메소드 추가
            samples.Types.Add(GameDataClass);

            // 각 테이블별 클래스 만들기
            for (int i = 0; i < dataSet.Tables.Count; i++)
            {
                CodeTypeDeclaration tempClass = new CodeTypeDeclaration(dataSet.Tables[i].TableName);
                tempClass.IsClass = true;
                tempClass.TypeAttributes = TypeAttributes.Public;
                tempClass.BaseTypes.Add("GameTable");

                for (int j = 0; j < dataSet.Tables[i].Columns.Count; j++)
                {
                    CodeMemberField memberProperty = new CodeMemberField(new CodeTypeReference(dataSet.Tables[i].Columns[j].DataType), AsteriskEraser(dataSet.Tables[i].Columns[j].ColumnName));
                    memberProperty.Attributes = MemberAttributes.Public | MemberAttributes.Final;
                    memberProperty.Name += " { get; private set; } //";
                    tempClass.Members.Add(memberProperty);
                }

                samples.Types.Add(tempClass);
            }

            CodeDomProvider provider = CodeDomProvider.CreateProvider("CSharp");
            CodeGeneratorOptions options = new CodeGeneratorOptions();
            options.BracingStyle = "C";
            using (StreamWriter sourceWriter = new StreamWriter(FilePath + @"/GameTable.cs"))
            {
                provider.GenerateCodeFromCompileUnit(targetUnit, sourceWriter, options);
            }
        }

        // * 표를 찾아서 지워준다.
        private static string AsteriskEraser(string name)
        {
            if (name.IndexOf("*", 0) == 0)
                name = name.Replace("*", string.Empty);

            return name;
        }
    }
}