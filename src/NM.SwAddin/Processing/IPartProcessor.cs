using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Processing
{
    public enum ProcessorType
    {
        SheetMetal,
        Tube,
        Machined,
        Generic
    }

    public sealed class ProcessingResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public NM.Core.ProblemParts.ProblemPartManager.ProblemCategory ErrorCategory { get; set; }
        public string ProcessorType { get; set; }
        public Dictionary<string, object> Data { get; } = new Dictionary<string, object>();

        public static ProcessingResult Ok(string type)
        {
            return new ProcessingResult { Success = true, ProcessorType = type };
        }
        public static ProcessingResult Fail(string msg, NM.Core.ProblemParts.ProblemPartManager.ProblemCategory cat)
        {
            return new ProcessingResult { Success = false, ErrorMessage = msg, ErrorCategory = cat };
        }
    }

    public interface IPartProcessor
    {
        ProcessorType Type { get; }
        bool CanProcess(IModelDoc2 model);
        ProcessingResult Process(IModelDoc2 model, NM.Core.ModelInfo info, NM.SwAddin.ProcessingOptions options);
    }
}
