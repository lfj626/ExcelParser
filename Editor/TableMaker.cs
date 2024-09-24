
using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using UnityEditor;

namespace MyTurn.ExcelParser
{
    public class TableMaker
    {
        #region Loader
        private static List<ExcelInfo> LoadTableList(string FilePath)
        {
            List<ExcelInfo> Infos = new List<ExcelInfo>();
            string TablePath = Path.Combine(FilePath, "TableList.xlsx");

            if (!File.Exists(TablePath))
                return Infos;
            
            using (var stream = File.Open(TablePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataset = reader.AsDataSet();
                    var table = dataset.Tables["TableList"];

                    for (int count = 0; count < table.Rows.Count; ++count)
                    {
                        // ���̺� �̸�, ���� �ǳʶٱ�
                        if (count <= 1)
                            continue;

                        // ��
                        if (table.Rows[count][1].ToString() == "@")
                            break;

                        // GSTR�� �н�
                        if (table.Rows[count][3].ToString().Contains("GSTR_"))
                            continue;

                        ExcelInfo NewInfo = new ExcelInfo()
                        {
                            FileName = table.Rows[count][2].ToString(),
                            TableName = table.Rows[count][3].ToString(),
                            ExportName = table.Rows[count][4].ToString(),
                            IsLocalTable = bool.Parse(table.Rows[count][5].ToString())
                        };

                        Infos.Add(NewInfo);
                    }
                }
            }

            return Infos;
        }

        private static List<ExcelInfo> LoadGSTRList(string FilePath)
        {
            List<ExcelInfo> Infos = new List<ExcelInfo>();
            string TablePath = Path.Combine(FilePath, "TableList.xlsx");

            if (!File.Exists(TablePath))
                return Infos;

            using (var stream = File.Open(TablePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataset = reader.AsDataSet();
                    var table = dataset.Tables["TableList"];

                    for (int count = 0; count < table.Rows.Count; ++count)
                    {
                        // ���̺� �̸�, ���� �ǳʶٱ�
                        if (count <= 1)
                            continue;

                        // ��
                        if (table.Rows[count][1].ToString() == "@")
                            break;

                        // GSTR�� �н�
                        if (table.Rows[count][3].ToString().Contains("GSTR_") == false)
                            continue;

                        ExcelInfo NewInfo = new ExcelInfo()
                        {
                            FileName = table.Rows[count][2].ToString(),
                            TableName = table.Rows[count][3].ToString(),
                            ExportName = table.Rows[count][4].ToString(),
                            IsLocalTable = bool.Parse(table.Rows[count][5].ToString())
                        };

                        Infos.Add(NewInfo);
                    }
                }
            }

            return Infos;
        }

        private static DataTable LoadTable(string FilePath, string FileName, string SheetName)
        {
            string TablePath = Path.Combine(FilePath, FileName);

            using (var stream = File.Open(TablePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataset = reader.AsDataSet();
                    var table = dataset.Tables[SheetName];

                    return OptimizeTable(table);
                }
            }
        }

        private static DataTable OptimizeTable(DataTable table)
        {
            // ����ȭ�� ���̺��� ����� �̸��� ������
            DataTable NewTable = new DataTable(table.Rows[0][1].ToString());

            // �÷������ - ������� �ʴ� �÷��� ���Խ�Ű�� ����
            for (int i = 1; i < table.Columns.Count; i++)
            {
                var data = table.Rows[2][i].ToString();
                if (data == "@")
                    break;

                // �÷� �̸�(Key)�� ������ �ڷ��� �־��ֱ�
                NewTable.Columns.Add(table.Rows[2][i].ToString(), DataTypeParser(table.Rows[3][i].ToString()));
            }

            // GSTR�� * �ٿ��ֱ� (�ڵ忡���� ����)
            for (int i = 1; i < table.Columns.Count; i++)
            {
                if (NewTable.Columns.Count <= i)
                    continue;

                if (table.Rows[3][i].ToString().Contains("GSTR"))
                {
                    NewTable.Columns[i - 1].ColumnName = string.Format("*{0}", NewTable.Columns[i - 1].ColumnName);
                }
            }

            // �ο��߰��ϱ�          
            for (int i = 0; i < table.Rows.Count - 5; i++)
            {
                DataRow NewRow = NewTable.NewRow();

                for (int j = 0; j < NewTable.Columns.Count; j++)
                {
                    if (table.Rows[i + 4][j + 1].ToString() == "@")
                        break;
                    try
                    {
                        NewRow[j] = table.Rows[i + 4][j + 1];
                    }
                    catch (Exception exception)
                    {
                        // TODO
                        break;
                    }
                }

                NewTable.Rows.Add(NewRow);
            }

            return NewTable;
        }
        #endregion

        #region Export
        public static async Task<bool> MakeTable(string TablePath, string FileSavePath)
        {
            try
            {
                var TableInfo = LoadTableList(TablePath);
                if (TableInfo == null)
                {
                    EditorUtility.DisplayDialog("Error", "Table List is empty", "");
                    return false;
                }

                if (Directory.Exists(FileSavePath) == false)
                    Directory.CreateDirectory(FileSavePath);

                foreach (var info in TableInfo)
                {
                    var Table = LoadTable(TablePath, info.FileName, info.TableName);
                    if (Table == null)
                    {
                        EditorUtility.DisplayDialog("Error", $"{info.FileName} does not have {info.TableName}", "");
                        return false;
                    }

                    string tableStr = TableToString(Table);
                    string SavePath = Path.Combine(FileSavePath, $"{Table.TableName}.bytes");
                    await File.WriteAllTextAsync(SavePath, tableStr);
                }

                string VersionData = Convert.ToBase64String(Encoding.UTF8.GetBytes(DateTime.Now.ToString("yyyyMMddHHmmss")));
                string VersionPath = Path.Combine(FileSavePath, "Version.bytes");
                await File.WriteAllTextAsync(VersionPath, VersionData);

                return true;
            }
            catch
            {
                return false;
            }
        }

        public static async Task<bool> MakeGSTR(string TablePath, string FileSavePath)
        {
            try
            {
                var TableInfo = LoadGSTRList(TablePath);
                if (TableInfo == null)
                {
                    EditorUtility.DisplayDialog("Error", "GSTR List is empty", "");
                    return false;
                }

                var GSTRInfo = GSTRToString(TablePath, TableInfo);
                if (GSTRInfo == null)
                {
                    EditorUtility.DisplayDialog("Error", "GSTR Export fail", "");
                    return false;
                }

                if (Directory.Exists(FileSavePath) == false)
                    Directory.CreateDirectory(FileSavePath);

                foreach (var Gstr in GSTRInfo.LocalInfo)
                {
                    string Lan = Gstr.Key;
                    string Data = JsonConvert.SerializeObject(Gstr.Value);
                    string Encrypt = Convert.ToBase64String(Encoding.UTF8.GetBytes(Data));

                    string FilePath = Path.Combine(FileSavePath, $"GSTR_Client_{Lan}.bytes");
                    await File.WriteAllTextAsync(FilePath, Encrypt);
                }

                foreach (var Gstr in GSTRInfo.PatchInfo)
                {
                    string Lan = Gstr.Key;
                    string Data = JsonConvert.SerializeObject(Gstr.Value);
                    string Encrypt = Convert.ToBase64String(Encoding.UTF8.GetBytes(Data));

                    string FilePath = Path.Combine(FileSavePath, $"GSTR_Bundle_{Lan}.bytes");
                    await File.WriteAllTextAsync(FilePath, Encrypt);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public static async Task<bool> MakeScript(string TablePath, string FileSavePath)
        {
            try
            {
                var TableInfo = LoadTableList(TablePath);
                if (TableInfo == null)
                {
                    EditorUtility.DisplayDialog("Error", "Table List is empty", "");
                    return false;
                }

                DataSet dataSet = new DataSet();
                foreach (var info in TableInfo)
                {
                    var Table = LoadTable(TablePath, info.FileName, info.TableName);
                    if (Table == null)
                    {
                        EditorUtility.DisplayDialog("Error", $"{info.FileName} does not have {info.TableName}", "");
                        return false;
                    }

                    dataSet.Tables.Add(Table);
                }

                ScriptMaker.MakeScript(FileSavePath, dataSet);
                await Task.Delay(100);

                return true;
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region Common
        private static Type DataTypeParser(string dataType) // OptimizeTable () ���� ����
        {
            switch (dataType)
            {
                case "I4":
                    return typeof(int);
                case "F4":
                    return typeof(float);
                case "STR":
                    return typeof(string);
                case "GSTR":
                    return typeof(string); // ���� ���������� �־���� �Ѵ�.
                case "BOOL":
                    return typeof(bool);
                default:
                    return null;
            }
        }

        private static string TableToString(DataTable Table)
        {
            Dictionary<int, JObject> RetValue = new Dictionary<int, JObject>();

            // �ο캰 �ڷ��� �̱�
            int Count = 0;
            foreach (DataRow row in Table.Rows)
            {
                JObject obj = new JObject();
                foreach (DataColumn col in Table.Columns)
                {
                    var value = row[col];
                    string pName = col.ColumnName.Replace("*", "");

                    obj.Add(pName, JToken.FromObject(value));
                }

                RetValue.Add(Count, obj);
                Count += 1;
            }

            string Json = JsonConvert.SerializeObject(RetValue);
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(Json));
        }

        private static GSTRInfo GSTRToString(string TablePath, List<ExcelInfo> Excels)
        {
            GSTRInfo TempData = new GSTRInfo();

            foreach (var excelInfo in Excels)
            {
                var CurTable = LoadTable(TablePath, excelInfo.FileName, excelInfo.TableName);
                if (CurTable == null)
                    continue;

                for (int colCount = 2; colCount < CurTable.Columns.Count; ++colCount)
                {
                    string colName = CurTable.Columns[colCount].ColumnName.ToUpper();
                    string Lang = colName.ToUpper();

                    for (int rowCount = 0; rowCount < CurTable.Rows.Count; ++rowCount)
                    {
                        string Key = CurTable.Rows[rowCount]["key"].ToString();
                        string Data = CurTable.Rows[rowCount][colName].ToString();
                        Data = Data.Replace("\\n", "\n");

                        if (excelInfo.IsLocalTable)
                        {
                            if (TempData.LocalInfo.ContainsKey(Lang))
                            {
                                if (TempData.LocalInfo[Lang] == null)
                                    TempData.LocalInfo[Lang] = new Dictionary<string, string>();

                                TempData.LocalInfo[Lang].Add(Key, Data);
                            }
                            else
                            {
                                TempData.LocalInfo.Add(Lang, new Dictionary<string, string>());
                                TempData.LocalInfo[Lang].Add(Key, Data);
                            }
                        }
                        else
                        {
                            if (TempData.PatchInfo.ContainsKey(Lang))
                            {
                                if (TempData.PatchInfo[Lang] == null)
                                    TempData.PatchInfo[Lang] = new Dictionary<string, string>();

                                TempData.PatchInfo[Lang].Add(Key, Data);
                            }
                            else
                            {
                                TempData.PatchInfo.Add(Lang, new Dictionary<string, string>());
                                TempData.PatchInfo[Lang].Add(Key, Data);
                            }
                        }
                    }
                }
            }

            return TempData;
        }
        #endregion
    }
}