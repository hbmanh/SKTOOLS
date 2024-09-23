namespace SKRevitAddins.Parameters
{
    internal class Define
    {
        // Default epsilon

        internal const double Epsilon = 1.0e-5;
        internal const double MinimumLength = 1 / 304.8;

        // Ribbon tab

        internal const string TNF_RibbonTab = "TBIM Tools";
        internal const string TNF_RibbonPanel = "積算";
        internal const string TNF_RibbonButton = "TNF";
        internal const string TNF_RibbonToolTip = "CSV出力";

        internal const string IconName = "TNF_Icon.png";
        internal const string IconSName = "TNF_Icon_S.png";
        internal const string DefaultNameOutput = "OutputFoundation.csv";
        internal const string TempCheckNameOutput = "TemplateCheck.csv";
        internal const string FilterFileCSV = "CSV files (*.csv)|*.csv";

        // Default name

        internal const string NameFamily = "土間コンクリート";
        internal const string Schedule_Export = "(積)";
        internal const string Table_Perimeter = "＜(積) アドイン集計_土間外周＞";
        internal const string Header_Perimeter1 = "タイプ名";
        internal const string Header_Perimeter2 = "隣接土間";
        internal const string Header_Perimeter3 = "厚さ(mm)";
        internal const string Header_Perimeter4 = "土間外周";
        internal const string Header_Perimeter5 = "段差コン";
        internal const string Header_Perimeter6 = "段差型枠";
        internal const string Header_Perimeter7 = "鉄筋_径";
        internal const string Header_Perimeter8 = "鉄筋_長さ";
        internal const string Header_Perimeter9 = "鉄筋_本数";
        internal const string Header_Perimeter10 = "鉄筋_ヶ所";
        internal const string Header_Perimeter11 = "鉄筋_全長";

        // New line default

        internal const string NewLineRevit = "\r\n";
        internal const string NewLineText = "\n";
        internal const string NextColumn = ",";

        // Message

        internal const string CaptionMessage = "積算";
        internal const string Message_ExportSuccess = "CSVファイルを出力しました。";
        internal const string Message_ExportFailed = "CSVファイルを出力することができませんでした。";
        internal const string Message_FileInUsing = "ファイルが開かれています。" + NewLineRevit + Message_ExportFailed;
        internal const string Message_NothingExport = "出力するデータがありません。" + NewLineRevit + Message_ExportFailed;
        internal const string Message_ErrorExport = "エラーが発生しました。" + NewLineRevit + Message_ExportFailed;
    }
}