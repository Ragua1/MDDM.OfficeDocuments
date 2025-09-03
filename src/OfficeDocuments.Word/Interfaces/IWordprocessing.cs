using System;

namespace OfficeDocuments.Word.Interfaces;

public interface IWordprocessing : IDisposable
{
    IBody GetBody();
    void Close(bool saveDocument = true);
}