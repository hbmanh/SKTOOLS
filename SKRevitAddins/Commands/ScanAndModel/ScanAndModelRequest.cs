using System.Threading;

namespace ScanAndModel
{
    public enum ScanAndModelRequestId
    {
        None = 0,
        AutoDetectAndModel,   // Tự động nhận diện & model
        ZoomToPoint           // Zoom đến vị trí point cloud để user tự model
    }

    public class ScanAndModelRequest
    {
        private int _request = (int)ScanAndModelRequestId.None;

        public void Make(ScanAndModelRequestId req)
        {
            Interlocked.Exchange(ref _request, (int)req);
        }

        public ScanAndModelRequestId Take()
        {
            return (ScanAndModelRequestId)Interlocked.Exchange(ref _request, (int)ScanAndModelRequestId.None);
        }
    }
}
