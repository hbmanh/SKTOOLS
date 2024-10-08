﻿using System.Threading;

namespace SKRevitAddins.Commands.PermissibleRangeFrame
{
    public enum RequestId : int
    {
        None = 0,
        OK = 1,
    }

    public class PermissibleRangeFrameRequest
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