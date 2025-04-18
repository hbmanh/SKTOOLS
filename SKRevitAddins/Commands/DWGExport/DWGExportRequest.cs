using System.Threading;

namespace SKRevitAddins.Commands.DWGExport
{
    public enum LayerExportRequestId
    {
        None = 0,
        Export
    }

    public class DWGExportRequest
    {
        private int _req = (int)LayerExportRequestId.None;

        public void Make(LayerExportRequestId r) => Interlocked.Exchange(ref _req, (int)r);
        public LayerExportRequestId Take() => (LayerExportRequestId)Interlocked.Exchange(ref _req, 0);
    }
}