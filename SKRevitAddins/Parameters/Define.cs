namespace SKRevitAddins.Parameters
{
    internal class Define
    {
        // Default epsilon

        internal const double Epsilon = 1.0e-5;
        internal const double MinimumLength = 1 / 304.8;

        // Ribbon tab

        internal const string TNF_RibbonTab = "TBIM Tools";
        internal const string TNF_RibbonPanel = "�ώZ";
        internal const string TNF_RibbonButton = "TNF";
        internal const string TNF_RibbonToolTip = "CSV�o��";

        internal const string IconName = "TNF_Icon.png";
        internal const string IconSName = "TNF_Icon_S.png";
        internal const string DefaultNameOutput = "OutputFoundation.csv";
        internal const string TempCheckNameOutput = "TemplateCheck.csv";
        internal const string FilterFileCSV = "CSV files (*.csv)|*.csv";

        // Default name

        internal const string NameFamily = "�y�ԃR���N���[�g";
        internal const string Schedule_Export = "(��)";
        internal const string Table_Perimeter = "��(��) �A�h�C���W�v_�y�ԊO����";
        internal const string Header_Perimeter1 = "�^�C�v��";
        internal const string Header_Perimeter2 = "�אړy��";
        internal const string Header_Perimeter3 = "����(mm)";
        internal const string Header_Perimeter4 = "�y�ԊO��";
        internal const string Header_Perimeter5 = "�i���R��";
        internal const string Header_Perimeter6 = "�i���^�g";
        internal const string Header_Perimeter7 = "�S��_�a";
        internal const string Header_Perimeter8 = "�S��_����";
        internal const string Header_Perimeter9 = "�S��_�{��";
        internal const string Header_Perimeter10 = "�S��_����";
        internal const string Header_Perimeter11 = "�S��_�S��";

        // New line default

        internal const string NewLineRevit = "\r\n";
        internal const string NewLineText = "\n";
        internal const string NextColumn = ",";

        // Message

        internal const string CaptionMessage = "�ώZ";
        internal const string Message_ExportSuccess = "CSV�t�@�C�����o�͂��܂����B";
        internal const string Message_ExportFailed = "CSV�t�@�C�����o�͂��邱�Ƃ��ł��܂���ł����B";
        internal const string Message_FileInUsing = "�t�@�C�����J����Ă��܂��B" + NewLineRevit + Message_ExportFailed;
        internal const string Message_NothingExport = "�o�͂���f�[�^������܂���B" + NewLineRevit + Message_ExportFailed;
        internal const string Message_ErrorExport = "�G���[���������܂����B" + NewLineRevit + Message_ExportFailed;
    }
}