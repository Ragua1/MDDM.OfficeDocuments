using System;

namespace OfficeDocumentsApi.Word.Interfaces
{
    public interface IWordprocessing : IDisposable
    {
        IBody AddBody();
    }
}