using System.Threading;

namespace SKRevitAddins.Commands.FindDWGNotUseAndDel
{
    public enum RequestId
    {
        None = 0,
        Delete,
        OpenView,
        Export,
    }

    public class FindDWGNotUseAndDelRequest
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
