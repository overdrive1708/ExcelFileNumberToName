namespace ExcelFileNumberToName.Models
{
    /// <summary>
    /// 調査設定クラス
    /// </summary>
    public class NumToNameSetting
    {
        /// <summary>
        /// 調査対象クラス
        /// </summary>
        public class ExaminationTarget
        {
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
        }
    }
}
