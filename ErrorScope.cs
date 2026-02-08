using System;

namespace NM.Core
{
    /// <summary>
    /// RAII wrapper for ErrorHandler.PushCallStack/PopCallStack pairs.
    /// Guarantees PopCallStack is called even on early returns or exceptions.
    ///
    /// Usage:
    ///   using (new ErrorScope("MyMethod"))
    ///   {
    ///       // code with call stack tracking
    ///   }
    /// </summary>
    public sealed class ErrorScope : IDisposable
    {
        private bool _disposed;

        public ErrorScope(string procedureName)
        {
            ErrorHandler.PushCallStack(procedureName);
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            ErrorHandler.PopCallStack();
        }
    }
}
