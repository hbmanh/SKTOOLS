using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ParamCopy
{
    public enum RequestId : int
    {
        None = 0,
        InstanceCopy = 1,
        FamilyCopy = 2,
        CategoryCopy = 3,
        AllEleCopy = 4
    }
    public class ParamCopyRequest
    {
        private int m_Request = (int)RequestId.None;

        public RequestId Take()
        {
            return (RequestId)Interlocked.Exchange(ref m_Request, (int)RequestId.None);
        }

        public void Make(RequestId request)
        {
            Interlocked.Exchange(ref m_Request, (int)request);
        }
    }
}
