using System;

namespace OfficeDocumentsApi.Word.Interfaces
{
    public interface IWordprocessing : IDisposable
    {
        IBody GetBody();
        void Close(bool saveDocument = true);
    }
}