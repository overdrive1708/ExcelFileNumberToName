using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelFileNumberToName.Models
{
    /// <summary>
    /// 調査ファイルクラス
    /// </summary>
    public class NumToNameFile
    {
        /// <summary>
        /// 調査結果クラス
        /// </summary>
        public class ExaminationResult
        {
            /// <summary>
            /// ファイル
            /// </summary>
            public string File { get; set; } = string.Empty;

            /// <summary>
            /// シート
            /// </summary>
            public string Sheet { get; set; } = string.Empty;

            /// <summary>
            /// セル
            /// </summary>
            public string Cell { get; set; } = string.Empty;

            /// <summary>
            /// メモ
            /// </summary>
            public string Memo { get; set; } = string.Empty;

            /// <summary>
            /// 数値
            /// </summary>
            public string Number { get; set; } = string.Empty;

            /// <summary>
            /// 名称
            /// </summary>
            public string Name { get; set; } = string.Empty;
        }

        /// <summary>
        /// 調査結果取得処理
        /// </summary>
        /// <param name="filename">ファイル名</param>
        /// <param name="examinationTargets">調査対象</param>
        /// <returns>調査結果</returns>
        public static List<ExaminationResult> GetExaminationResult(string filename, List<NumToNameSetting.ExaminationTarget> examinationTargets)
        {
            List<ExaminationResult> results = [];
            ExaminationResult result;

            // Bookを開く(読み取り専用で開く)
            using (FileStream fs = new(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using XLWorkbook workbook = new(fs);

                // 調査対象で指定されたセルの調査
                foreach (NumToNameSetting.ExaminationTarget examinationTarget in examinationTargets)
                {
                    // ワークシートを開く
                    if (workbook.TryGetWorksheet(examinationTarget.Sheet, out IXLWorksheet worksheet))
                    {
                        if (!examinationTarget.Cell.Contains(':'))
                        {
                            // 単一セルの場合
                            MatchCollection regexMatchResults = Regex.Matches(worksheet.Cell(examinationTarget.Cell).Value.ToString(), @"[0-9]+");
                            foreach (Match regexMatchResult in regexMatchResults.Cast<Match>())
                            {
                                // 正規表現で数値を抽出して変換結果を結果とする
                                result = new()
                                {
                                    File = filename,
                                    Sheet = examinationTarget.Sheet,
                                    Cell = examinationTarget.Cell,
                                    Memo = examinationTarget.Memo,
                                    Number = regexMatchResult.Value,
                                    Name = NumToNameRule.GetName(regexMatchResult.Value)
                                };
                                results.Add(result);
                            }
                        }
                        else
                        {
                            // 複数セルの場合
                            // 範囲指定でセルの値を取得
                            IXLRange table = worksheet.Range(examinationTarget.Cell).AsTable();
                            foreach (IXLRangeRow rowData in table.Rows())
                            {
                                foreach (IXLCell cellData in rowData.Cells())
                                {
                                    MatchCollection regexMatchResults = Regex.Matches(cellData.Value.ToString(), @"[0-9]+");
                                    foreach (Match regexMatchResult in regexMatchResults.Cast<Match>())
                                    {
                                        // 正規表現で数値を抽出して変換結果を結果とする
                                        result = new()
                                        {
                                            File = filename,
                                            Sheet = examinationTarget.Sheet,
                                            Cell = examinationTarget.Cell,
                                            Memo = examinationTarget.Memo,
                                            Number = regexMatchResult.Value,
                                            Name = NumToNameRule.GetName(regexMatchResult.Value)
                                        };
                                        results.Add(result);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return results;
        }
    }
}
