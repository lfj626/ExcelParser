using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using Newtonsoft.Json;


namespace MyTurn.ExcelParser
{
    class ExcelParserPath
    {
        public string TablePath { get; set; }
        public string ScriptOutput { get; set; }
        public string TableOutput { get; set; }
        public string GSTROutput { get; set; }
        public bool IsSecure { get; set; }
        public string SecureKey { get; set; }
    }

    public class ExcelInfo
    {
        public string FileName { get; set; }
        public string TableName { get; set; }
        public string ExportName { get; set; }
        public bool IsLocalTable { get; set; }
    }

    public class GSTRInfo
    {
        public Dictionary<string, Dictionary<string, string>> LocalInfo = new Dictionary<string, Dictionary<string, string>>();
        public Dictionary<string, Dictionary<string, string>> PatchInfo = new Dictionary<string, Dictionary<string, string>>();
    }

    public class ExcelParserWindow : EditorWindow
    {
        private GUILayoutOption[] LabelOption = { GUILayout.MinWidth(10f), GUILayout.MaxWidth(100) };
        private GUILayoutOption[] TextOption = { GUILayout.MinWidth(10f), GUILayout.MaxWidth(250f) };
        private GUILayoutOption[] OtherOption = { GUILayout.MinWidth(10f), GUILayout.MaxWidth(150f) };
        private GUILayoutOption[] ButtonOption = { GUILayout.MinWidth(10f), GUILayout.MaxWidth(250f) };
        private Vector2 ScrollBar;

        private ExcelParserPath PathData = null;

        [MenuItem("Tool/ExcelParser")]
        private static void OpenWindow()
        {
            var Window = GetWindow<ExcelParserWindow>();
            Window.LoadWindow();
        }

        private void OnEnable()
        {
            LoadWindow();
        }

        private void LoadWindow()
        {
            if (PathData == null)
            {
                string SaveData = EditorPrefs.GetString(Application.productName, string.Empty);
                if (string.IsNullOrEmpty(SaveData))
                    PathData = new ExcelParserPath();
                else
                    PathData = JsonConvert.DeserializeObject<ExcelParserPath>(SaveData);
            }
        }

        private void SaveWindow()
        {
            var Json = JsonConvert.SerializeObject(PathData);
            EditorPrefs.SetString(Application.productName, Json);
        }

        private void OnClick_AllExport()
        {
            if (PathData == null)
            {
                EditorUtility.DisplayDialog("Error", "Path Invalid", "OK");
                return;
            }

            if (string.IsNullOrEmpty(PathData.TablePath))
            {
                EditorUtility.DisplayDialog("Error", "Table Path is empty", "OK");
                return;
            }

            if (TableMaker.MakeTable(PathData.TablePath, PathData.TableOutput) == false)
            {
                EditorUtility.DisplayDialog("Error", "Table Export Error", "OK");
                return;
            }

            if (TableMaker.MakeGSTR(PathData.TablePath, PathData.GSTROutput) == false)
            {
                EditorUtility.DisplayDialog("Error", "GSTR Export Error", "OK");
                return;
            }

            if (TableMaker.MakeScript(PathData.TablePath, PathData.ScriptOutput) == false)
            {
                EditorUtility.DisplayDialog("Error", "Script Export Error", "OK");
                return;
            }

            AssetDatabase.Refresh();

            EditorUtility.DisplayDialog("Success", "Done", "OK");
        }

        private void OnClick_ScriptExport()
        {
            if (PathData == null)
            {
                EditorUtility.DisplayDialog("Error", "Path Invalid", "OK");
                return;
            }

            if (string.IsNullOrEmpty(PathData.ScriptOutput))
            {
                EditorUtility.DisplayDialog("Error", "Script path is empty", "OK");
                return;
            }

            if (TableMaker.MakeScript(PathData.TablePath, PathData.ScriptOutput) == false)
            {
                EditorUtility.DisplayDialog("Error", "Script Export Error", "OK");
                return;
            }

            AssetDatabase.Refresh();

            EditorUtility.DisplayDialog("Success", "Done", "OK");
        }

        private void OnClick_TableExport()
        {
            if (PathData == null)
            {
                EditorUtility.DisplayDialog("Error", "Path Invalid", "OK");
                return;
            }

            if (string.IsNullOrEmpty(PathData.TableOutput))
            {
                EditorUtility.DisplayDialog("Error", "Table path is empty", "OK");
                return;
            }

            if (TableMaker.MakeTable(PathData.TablePath, PathData.TableOutput) == false)
            {
                EditorUtility.DisplayDialog("Error", "Table Export Error", "OK");
                return;
            }

            AssetDatabase.Refresh();

            EditorUtility.DisplayDialog("Success", "Done", "OK");
        }

        private void OnClick_GSTRExport()
        {
            if (PathData == null)
            {
                EditorUtility.DisplayDialog("Error", "Path Invalid", "OK");
                return;
            }

            if (string.IsNullOrEmpty(PathData.GSTROutput))
            {
                EditorUtility.DisplayDialog("Error", "GSTR path is empty", "OK");
                return;
            }

            if (TableMaker.MakeGSTR(PathData.TablePath, PathData.GSTROutput) == false)
            {
                EditorUtility.DisplayDialog("Error", "GSTR Export Error", "OK");
                return;
            }

            AssetDatabase.Refresh();

            EditorUtility.DisplayDialog("Success", "Done", "OK");
        }

        private void OnGUI()
        {
            ScrollBar = GUILayout.BeginScrollView(ScrollBar);

            Display();

            GUILayout.EndScrollView();
        }

        private void Display()
        {
            // 테이블 경로
            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField("Excel Path", EditorStyles.label, LabelOption);
            EditorGUILayout.TextArea(PathData.TablePath, EditorStyles.textArea, TextOption);
            if (GUILayout.Button("Select", OtherOption))
            {
                PathData.TablePath = EditorUtility.SaveFolderPanel("Choose folder", "", "");

                if (string.IsNullOrEmpty(PathData.TablePath) == false)
                    SaveWindow();
            }
            EditorGUILayout.EndHorizontal();

            // 스크립트 경로
            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField("Script Output", EditorStyles.label, LabelOption);
            EditorGUILayout.TextArea(PathData.ScriptOutput, EditorStyles.textArea, TextOption);
            if (GUILayout.Button("Select", OtherOption))
            {
                PathData.ScriptOutput = EditorUtility.SaveFolderPanel("Choose folder", "", "");

                if (string.IsNullOrEmpty(PathData.ScriptOutput) == false)
                    SaveWindow();
            }
            EditorGUILayout.EndHorizontal();

            // 테이블 추출 경로
            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField("Table Output", EditorStyles.label, LabelOption);
            EditorGUILayout.TextArea(PathData.TableOutput, EditorStyles.textArea, TextOption);
            if (GUILayout.Button("Select", OtherOption))
            {
                PathData.TableOutput = EditorUtility.SaveFolderPanel("Choose folder", "", "");

                if (string.IsNullOrEmpty(PathData.TableOutput) == false)
                    SaveWindow();
            }
            EditorGUILayout.EndHorizontal();

            // GSTR 추출 경로
            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField("GSTR Output", EditorStyles.label, LabelOption);
            EditorGUILayout.TextArea(PathData.GSTROutput, EditorStyles.textArea, TextOption);
            if (GUILayout.Button("Select", OtherOption))
            {
                PathData.GSTROutput = EditorUtility.SaveFolderPanel("Choose folder", "", "");

                if (string.IsNullOrEmpty(PathData.GSTROutput) == false)
                    SaveWindow();
            }
            EditorGUILayout.EndHorizontal();

            // 암호화
            EditorGUILayout.BeginHorizontal();
            EditorGUILayout.LabelField("Secure Mode", EditorStyles.label, LabelOption);
            bool IsSecure = EditorGUILayout.Toggle(PathData.IsSecure, EditorStyles.toggle, TextOption);
            if (PathData.IsSecure != IsSecure)
            {
                PathData.IsSecure = IsSecure;
                SaveWindow();
            }
            EditorGUILayout.EndHorizontal();

            if (PathData.IsSecure)
            {
                EditorGUILayout.BeginHorizontal();
                EditorGUILayout.LabelField("Secure Key", EditorStyles.label, LabelOption);
                PathData.SecureKey = EditorGUILayout.TextArea(PathData.SecureKey, EditorStyles.textArea, TextOption);
                if (GUILayout.Button("Save", OtherOption))
                {
                    SaveWindow();
                }
                EditorGUILayout.EndHorizontal();
            }

            EditorGUILayout.Space();

            // 버튼
            EditorGUILayout.BeginHorizontal();
            if (GUILayout.Button("All", ButtonOption))
                OnClick_AllExport();

            if (GUILayout.Button("Script", ButtonOption))
                OnClick_ScriptExport();
            EditorGUILayout.EndHorizontal();

            EditorGUILayout.BeginHorizontal();
            if (GUILayout.Button("Table", ButtonOption))
                OnClick_TableExport();

            if (GUILayout.Button("GSTR", ButtonOption))
                OnClick_GSTRExport();
            EditorGUILayout.EndHorizontal();
        }
    }
}