﻿using ExcelFileNumberToName.Models;
using Prism.Commands;
using Prism.Mvvm;
using System.Collections.ObjectModel;
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

        }

        /// <summary>
        /// プリセット選択変更コマンド実行処理
        /// </summary>
        private void ExecuteCommandPresetChange()
        {

        }

        /// <summary>
        /// 調査ファイルドラッグコマンド実行処理
        /// </summary>
        /// <param name="e">イベントデータ</param>
        private void ExecuteCommandExaminationFilePreviewDragOver(DragEventArgs e)
        {

        }

        /// <summary>
        /// 調査ファイルドロップコマンド実行処理
        /// </summary>
        /// <param name="e">イベントデータ</param>
        private void ExecuteCommandExaminationFileDrop(DragEventArgs e)
        {

        }

        /// <summary>
        /// 調査ファイルクリアコマンド実行処理
        /// </summary>
        private void ExecuteCommandClearExaminationFile()
        {

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

        }
    }
}
