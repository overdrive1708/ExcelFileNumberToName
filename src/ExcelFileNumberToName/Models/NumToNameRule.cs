using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ExcelFileNumberToName.Models
{
    /// <summary>
    /// 変換ルールクラス
    /// </summary>
    public class NumToNameRule
    {
        /// <summary>
        /// ルールクラス
        /// </summary>
        public class Rule
        {
            /// <summary>
            /// 数値
            /// </summary>
            public string Number { get; set; } = string.Empty;

            /// <summary>
            /// 名称
            /// </summary>
            public string Name { get; set; } = string.Empty;
        }

        //--------------------------------------------------
        // 定数(コンフィギュレーション)
        //--------------------------------------------------
        /// <summary>
        /// 変換ルールファイル名
        /// </summary>
        private static readonly string _fileName = "ConvertRules.csv";

        //--------------------------------------------------
        // 内部変数
        //--------------------------------------------------
        /// <summary>
        /// 変換ルール
        /// </summary>
        private static List<Rule> _rules = [];

        //--------------------------------------------------
        // メソッド
        //--------------------------------------------------
        /// <summary>
        /// 変換ルールファイル読み込み処理
        /// </summary>
        public static void Read()
        {
            // 設定ファイルパス作成
            string settingFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _fileName);

            // 設定ファイルがない場合は新規作成する
            if (!File.Exists(settingFilePath))
            {
                Write();
            }

            // 設定ファイルの読み込み
            CsvConfiguration options = new(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true
            };

            using StreamReader srCsv = new(settingFilePath, System.Text.Encoding.UTF8);
            using CsvHelper.CsvReader csv = new(srCsv, options);
            _rules = csv.GetRecords<Rule>().ToList();

            srCsv.Close();
        }

        /// <summary>
        /// 変換ルールファイル書き込み処理
        /// </summary>
        private static void Write()
        {
            // 設定ファイルパス作成
            string settingFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _fileName);

            using StreamWriter swCsv = new(settingFilePath, false, System.Text.Encoding.UTF8);

            using (CsvHelper.CsvWriter csv = new(swCsv, CultureInfo.InvariantCulture))
            {
                csv.WriteHeader<Rule>();
                csv.NextRecord();
                csv.WriteRecords(_rules);
            }

            swCsv.Close();
        }

        /// <summary>
        /// 名称取得処理
        /// </summary>
        /// <param name="number">数値</param>
        /// <returns></returns>
        public static string GetName(string number)
        {
            string name = string.Empty;     // 見つからない場合は空

            // 変換ルールを検索して数値を名称に変換する
            foreach (var rule in from Rule rule in _rules
                                 where rule.Number == number
                                 select rule)
            {
                name = rule.Name;
            }

            return name;
        }
    }
}
