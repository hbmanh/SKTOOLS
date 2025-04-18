using System.IO;

namespace SKRevitAddins.Commands.LayoutsToDWG
{
    /// <summary>
    /// Đường dẫn file script .scr cho AutoCAD
    /// </summary>
    public static class AcadAppFolder
    {
        public static string AcadCmdFile => Path.Combine(Path.GetTempPath(), "acadCommand.scr");
    }
}
