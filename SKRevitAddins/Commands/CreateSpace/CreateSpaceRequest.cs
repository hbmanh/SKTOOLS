using System.Threading;

namespace SKRevitAddins.Commands.CreateSpace
{
    public enum RequestId : int
    {
        None = 0,
        CreateSpace = 1,
        DeleteSpace = 2,
    }

    public class CreateSpaceRequest
    {
        private int m_request = (int)RequestId.None;

        public RequestId Take()
        {
            return (RequestId)Interlocked.Exchange(ref m_request, (int)RequestId.None);
        }

        public void Make(RequestId request)
        {
            Interlocked.Exchange(ref m_request, (int)request);
        }
    }
}
