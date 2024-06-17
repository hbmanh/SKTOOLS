using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SKToolsAddins.Commands.ChangeBwTypeAndIns
{
    public enum RequestId : int
    {
        None = 0,
        OK = 1,
    }

    public class ChangeBwTypeAndInsRequest
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