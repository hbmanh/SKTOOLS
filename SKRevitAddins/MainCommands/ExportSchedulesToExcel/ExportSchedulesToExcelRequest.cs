using System.Threading;

namespace SKRevitAddins.ExportSchedulesToExcel
{
    public enum RequestId
    {
        None = 0,
        Export
    }

    public class ExportSchedulesToExcelRequest
    {
        private int _request = (int)RequestId.None;

        public void Make(RequestId req)
        {
            Interlocked.Exchange(ref _request, (int)req);
        }

        public RequestId Take()
        {
            return (RequestId)Interlocked.Exchange(ref _request, (int)RequestId.None);
        }
    }
}
