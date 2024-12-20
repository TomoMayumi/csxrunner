using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using System.Threading.Tasks;

namespace csxrunner
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            LoadScripts();
        }
        private void LoadScripts()
        {
            string scriptFolder = @"c:\work\cs\csx";
            var files = System.IO.Directory.GetFiles(scriptFolder, "*.csx");

            scriptComboBox.Items.Clear(); // 既存のアイテムをクリア

            foreach (var file in files)
            {
                string fileName = System.IO.Path.GetFileName(file);
                var item = this.Factory.CreateRibbonDropDownItem();
                item.Label = fileName; // ドロップダウンアイテムのラベルを設定
                scriptComboBox.Items.Add(item); // アイテムを追加
            }
        }
        private async void RunButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (scriptComboBox.SelectedItem != null)
            {
                string selectedScript = scriptComboBox.SelectedItem.ToString();
                string scriptPath = System.IO.Path.Combine(@"c:\work\cs\csx", selectedScript);

                // C#スクリプトを実行
                await ExecuteScript(scriptPath);
            }
        }

        private async Task ExecuteScript(string scriptPath)
        {
            try
            {
                // スクリプトファイルの内容を読み込む
                string scriptCode = System.IO.File.ReadAllText(scriptPath);

                // スクリプトを実行する
                var result = await CSharpScript.EvaluateAsync(scriptCode, ScriptOptions.Default);

                // 結果を表示する（必要に応じて）
                System.Windows.Forms.MessageBox.Show(result?.ToString() ?? "スクリプトが正常に実行されました。");
            }
            catch (Exception ex)
            {
                // エラーハンドリング
                System.Windows.Forms.MessageBox.Show($"エラー: {ex.Message}");
            }
        }
    }
}
