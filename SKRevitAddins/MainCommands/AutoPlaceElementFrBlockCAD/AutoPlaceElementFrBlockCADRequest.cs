using System.Threading;

namespace SKRevitAddins.AutoPlaceElementFrBlockCAD
{
    public enum RequestId : int
    {
        None = 0,
        OK = 1,
    }

    public class AutoPlaceElementFrBlockCADRequest
    {
        private int m_request = (int)RequestId.None;

        public RequestId Take()
        {
            return (RequestId)System.Threading.Interlocked.Exchange(ref m_request, (int)RequestId.None);
        }

        public void Make(RequestId request)
        {
            System.Threading.Interlocked.Exchange(ref m_request, (int)request);
        }
    }
}
