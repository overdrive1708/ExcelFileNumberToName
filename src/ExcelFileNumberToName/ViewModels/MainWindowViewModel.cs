using ExcelFileNumberToName.Models;
using Prism.Commands;
using Prism.Mvvm;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelFileNumberToName.ViewModels
{
    public class MainWindowViewModel : BindableBase
    {
        //--------------------------------------------------
        // クラス
        //--------------------------------------------------
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

        //--------------------------------------------------
        // バインディングデータ
        //--------------------------------------------------
        /// <summary>
        /// タイトル
        /// </summary>
        private string _title = Resources.Strings.ApplicationName;
        public string Title
        {
            get { return _title; }
            set { SetProperty(ref _title, value); }
        }

        /// <summary>
        /// 調査設定：プリセット(リスト)
        /// </summary>
        private ObservableCollection<string> _presetList = [];
        public ObservableCollection<string> PresetList
        {
            get { return _presetList; }
            set { SetProperty(ref _presetList, value); }
        }

        /// <summary>
        /// 調査設定：プリセット(選択状態)
        /// </summary>
        private string _preset = string.Empty;
        public string Preset
        {
            get { return _preset; }
            set { SetProperty(ref _preset, value); }
        }

        /// <summary>
        /// 調査設定：調査ファイルキーワード
        /// </summary>
        private string _examinationFileKeyword = string.Empty;
        public string ExaminationFileKeyword
        {
            get { return _examinationFileKeyword; }
            set { SetProperty(ref _examinationFileKeyword, value); }
        }

        /// <summary>
        /// 調査設定：調査対象(リスト)
        /// </summary>
        private ObservableCollection<NumToNameSetting.ExaminationTarget> _examinationTargetList = [];
        public ObservableCollection<NumToNameSetting.ExaminationTarget> ExaminationTargetList
        {
            get { return _examinationTargetList; }
            set { SetProperty(ref _examinationTargetList, value); }
        }

        /// <summary>
        /// 調査ファイル(リスト)
        /// </summary>
        private ObservableCollection<string> _examinationFileList = [];
        public ObservableCollection<string> ExaminationFileList
        {
            get { return _examinationFileList; }
            set { SetProperty(ref _examinationFileList, value); }
        }

        /// <summary>
        /// 調査ファイルガイド表示可否
        /// </summary>
        private bool _examinationFileGuideVisibility = true;
        public bool ExaminationFileGuideVisibility
        {
            get { return _examinationFileGuideVisibility; }
            set { SetProperty(ref _examinationFileGuideVisibility, value); }
        }

        /// <summary>
        /// 操作可能フラグ
        /// </summary>
        private bool _isOperationEnable = true;
        public bool IsOperationEnable
        {
            get { return _isOperationEnable; }
            set { SetProperty(ref _isOperationEnable, value); }
        }

        /// <summary>
        /// 調査結果(リスト)
        /// </summary>
        private ObservableCollection<ExaminationResult> _examinationResultList = [];
        public ObservableCollection<ExaminationResult> ExaminationResultList
        {
            get { return _examinationResultList; }
            set { SetProperty(ref _examinationResultList, value); }
        }

        /// <summary>
        /// コピーライト
        /// </summary>
        private string _copyright = string.Empty;
        public string Copyright
        {
            get { return _copyright; }
            set { SetProperty(ref _copyright, value); }
        }

        /// <summary>
        /// プログレスバー最大値
        /// </summary>
        private int _progressMaximum = 1;
        public int ProgressMaximum
        {
            get { return _progressMaximum; }
            set { SetProperty(ref _progressMaximum, value); }
        }

        /// <summary>
        /// プログレスバー現在値
        /// </summary>
        private int _progressValue = 0;
        public int ProgressValue
        {
            get { return _progressValue; }
            set { SetProperty(ref _progressValue, value); }
        }

        /// <summary>
        /// 進捗メッセージ
        /// </summary>
        private string _progressMessage = string.Empty;
        public string ProgressMessage
        {
            get { return _progressMessage; }
            set { SetProperty(ref _progressMessage, value); }
        }

        //--------------------------------------------------
        // バインディングコマンド
        //--------------------------------------------------
        /// <summary>
        /// プリセット選択変更
        /// </summary>
        private DelegateCommand _commandPresetChange;
        public DelegateCommand CommandPresetChange => _commandPresetChange ??= new DelegateCommand(ExecuteCommandPresetChange);

        /// <summary>
        /// 調査ファイルドラッグ
        /// </summary>
        private DelegateCommand<DragEventArgs> _commandExaminationFilePreviewDragOver;
        public DelegateCommand<DragEventArgs> CommandExaminationFilePreviewDragOver => _commandExaminationFilePreviewDragOver ??= new DelegateCommand<DragEventArgs>(ExecuteCommandExaminationFilePreviewDragOver);

        /// <summary>
        /// 調査ファイルドロップ
        /// </summary>
        private DelegateCommand<DragEventArgs> _commandExaminationFileDrop;
        public DelegateCommand<DragEventArgs> CommandExaminationFileDrop => _commandExaminationFileDrop ??= new DelegateCommand<DragEventArgs>(ExecuteCommandExaminationFileDrop);

        /// <summary>
        /// 調査ファイルクリア
        /// </summary>
        private DelegateCommand _commandClearExaminationFile;
        public DelegateCommand CommandClearExaminationFile => _commandClearExaminationFile ??= new DelegateCommand(ExecuteCommandClearExaminationFile);

        /// <summary>
        /// 調査実施
        /// </summary>
        private DelegateCommand _commandExamination;
        public DelegateCommand CommandExamination => _commandExamination ??= new DelegateCommand(ExecuteCommandExamination);

        /// <summary>
        /// URLを開く
        /// </summary>
        private DelegateCommand<string> _commandOpenUrl;
        public DelegateCommand<string> CommandOpenUrl => _commandOpenUrl ??= new DelegateCommand<string>(ExecuteCommandOpenUrl);

        //--------------------------------------------------
        // メソッド
        //--------------------------------------------------
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public MainWindowViewModel()
        {
            Assembly assm = Assembly.GetExecutingAssembly();
            string version = assm.GetCustomAttribute<AssemblyInformationalVersionAttribute>().InformationalVersion;

            // バージョン情報を取得してタイトルに反映する
            Title = $"{Resources.Strings.ApplicationName} Ver.{version}";

            // コピーライト情報を取得して設定する
            Copyright = assm.GetCustomAttribute<AssemblyCopyrightAttribute>().Copyright;

            // 設定ファイルの読み込み
            // (XAMLデザイナーのエラー対策でデザインモードではない場合のみ)
            if (!(bool)System.ComponentModel.DesignerProperties.IsInDesignModeProperty.GetMetadata(typeof(DependencyObject)).DefaultValue)
            {
                NumToNameSetting.Read();
            }

            // プリセット(リスト)の反映
            PresetList = new(NumToNameSetting.GetPresetList());

            // 調査設定の反映(プリセットリスト先頭)
            if (PresetList.Count != 0)
            {
                LoadExaminationSettings(PresetList[0]);
            }

            // 調査実施できるか確認
            CheckExecuteExamination();
        }

        /// <summary>
        /// プリセット選択変更コマンド実行処理
        /// </summary>
        private void ExecuteCommandPresetChange()
        {
            // 調査設定の反映(選択したプリセットリスト)
            LoadExaminationSettings(Preset);
        }

        /// <summary>
        /// 調査ファイルドラッグコマンド実行処理
        /// </summary>
        /// <param name="e">イベントデータ</param>
        private void ExecuteCommandExaminationFilePreviewDragOver(DragEventArgs e)
        {
            // ドラッグしてきたデータがファイルの場合､ドロップを許可する｡
            e.Effects = DragDropEffects.Copy;
            e.Handled = e.Data.GetDataPresent(DataFormats.FileDrop);
        }

        /// <summary>
        /// 調査ファイルドロップコマンド実行処理
        /// </summary>
        /// <param name="e">イベントデータ</param>
        private void ExecuteCommandExaminationFileDrop(DragEventArgs e)
        {
            // ドロップしてきたデータを解析する
            if (e.Data.GetData(DataFormats.FileDrop) is string[] dropitems)
            {
                foreach (string dropitem in dropitems)
                {
                    if (System.IO.Directory.Exists(dropitem))
                    {
                        // フォルダの場合は配下の調査ファイルキーワードを含むサポートExcelファイルを調査ファイルのリストに追加
                        if (System.IO.Directory.GetFiles(@dropitem, "*", System.IO.SearchOption.AllDirectories) is string[] files)
                        {
                            foreach (string file in files)
                            {
                                if (System.IO.Path.GetFileName(file).Contains(ExaminationFileKeyword) && (System.IO.Path.GetExtension(file) == ".xlsx" || System.IO.Path.GetExtension(file) == ".xlsm"))
                                {
                                    ExaminationFileList.Add(file);
                                }
                            }
                        }
                    }
                    else
                    {
                        // ファイルの場合は調査ファイルキーワードを含むサポートExcelファイルを調査ファイルのリストに追加
                        if (System.IO.Path.GetFileName(dropitem).Contains(ExaminationFileKeyword) && (System.IO.Path.GetExtension(dropitem) == ".xlsx" || System.IO.Path.GetExtension(dropitem) == ".xlsm"))
                        {
                            ExaminationFileList.Add(dropitem);
                        }
                    }
                }
            }

            // 調査実施できるか確認
            CheckExecuteExamination();
        }

        /// <summary>
        /// 調査ファイルクリアコマンド実行処理
        /// </summary>
        private void ExecuteCommandClearExaminationFile()
        {
            // 調査ファイルリストをクリア
            ExaminationFileList.Clear();
            ExaminationFileGuideVisibility = true;

            // 調査実施できるか確認
            CheckExecuteExamination();
        }

        /// <summary>
        /// 調査実施コマンド実行処理
        /// </summary>
        private async void ExecuteCommandExamination()
        {
            await Task.Run(() =>
            {

            });
        }

        /// <summary>
        /// URLを開くコマンド実行処理
        /// </summary>
        /// <param name="url">URL</param>
        private void ExecuteCommandOpenUrl(string url)
        {
            ProcessStartInfo psi = new() { FileName = url, UseShellExecute = true };
            _ = Process.Start(psi);
        }

        /// <summary>
        /// 調査設定反映処理
        /// </summary>
        /// <param name="presetName">プリセット名</param>
        private void LoadExaminationSettings(string presetName)
        {
            // プリセット取得
            NumToNameSetting.Setting setting = NumToNameSetting.GetPreset(presetName);

            // プリセット(選択状態)に反映
            Preset = presetName;

            // 調査ファイルキーワードに反映
            ExaminationFileKeyword = setting.ExaminationFileKeyword;

            // 調査対象(リスト)に反映
            ExaminationTargetList = new(setting.ExaminationTargets);

            // 調査実施できるか確認
            CheckExecuteExamination();
        }

        /// <summary>
        /// 調査実施可否確認処理
        /// </summary>
        private void CheckExecuteExamination()
        {
            if (ExaminationTargetList.Count == 0)
            {
                // 調査対象の設定がない場合は調査実施不可
                IsOperationEnable = false;
                ProgressMessage = Resources.Strings.MessageStatusExaminationTargetEmpty;
            }
            else if (ExaminationFileList.Count == 0)
            {
                // 調査ファイルがない場合は調査実施不可
                IsOperationEnable = false;
                ProgressMessage = Resources.Strings.MessageStatusExaminationFileEmpty;
            }
            else
            {
                // チェック通過時調査実施可能
                IsOperationEnable = true;
                ProgressMessage = Resources.Strings.MessageStatusAlreadyExamination;
                ExaminationFileGuideVisibility = false;
            }
        }
    }
}
