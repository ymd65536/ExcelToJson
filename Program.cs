using System;
using System.IO;
using ClosedXML.Excel;

namespace ExcelToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            string ConfigPath = args[0];
            string ExcelPath = args[1];
            string KeyName = args[2];
            string OutputFolder = args[3];

            //ワイルドカード"*"は、すべてのファイルを意味する
            string[] configs = Directory.GetFiles(
                ConfigPath, "*", SearchOption.AllDirectories);

            string ConfigFile = configs[0];
            string[] values;
            string OutputJson;


            if (File.Exists(configs[0]))
            {
                string[] ConfigText = File.ReadAllLines(ConfigFile);
                //ワイルドカード"*"は、すべてのファイルを意味する
                string[] files = Directory.GetFiles(
                    ExcelPath, "*", SearchOption.TopDirectoryOnly);

                if (Directory.Exists(OutputFolder))
                {
                    Directory.Delete(OutputFolder, true);
                }
                Directory.CreateDirectory(OutputFolder);

                foreach (string file in files)
                {
                    values = new string[ConfigText.Length];
                    int cnt = 0;
                    foreach (string Line in ConfigText)
                    {
                        XLWorkbook workbook = new(file);
                        IXLWorksheet worksheet = workbook.Worksheet(1);

                        // Excelから指定の箇所を取得
                        string Rng = worksheet.Cell(Line).Value.ToString();
                        values[cnt] = "\"" + Line + "\":\"" + Rng + "\"";
                        cnt++;
                    }

                    // json ファイルに出力
                    OutputJson = "{\"" + KeyName + "\":{";
                    string JsonValue = String.Join(",", values);
                    OutputJson = OutputJson + JsonValue + "}}";

                    values = null;
                    string JsonFileName = Path.GetFileName(file).Replace(".xlsx", "");

                    using (StreamWriter sw = new StreamWriter(OutputFolder + "/" + JsonFileName + ".json"))
                    {
                        sw.WriteLine(OutputJson);
                    }
                }

            }
            else
            {
                return;
            }
        }
    }
}
