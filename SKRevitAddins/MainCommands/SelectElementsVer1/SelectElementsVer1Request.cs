﻿using System.Threading;

namespace SKRevitAddins.SelectElementsVer1
{
    public enum RequestId : int
    {
        None = 0,
        OK = 1,
    }

    public class SelectElementsVer1Request
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