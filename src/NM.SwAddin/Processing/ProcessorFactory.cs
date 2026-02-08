using NM.Core;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin.Processing
{
    public interface IProcessorFactory
    {
        IPartProcessor Get(ProcessorType type);
        IPartProcessor DetectFor(IModelDoc2 model);
    }

    public sealed class ProcessorFactory : IProcessorFactory
    {
        private readonly ISldWorks _swApp;
        private readonly IPartProcessor _generic;
        private readonly IPartProcessor _sheet;
        private readonly IPartProcessor _tube;
        private readonly IPartProcessor _machined;

        /// <summary>
        /// Convenience constructor that creates all processors with ISldWorks dependency.
        /// This is the preferred constructor for production use.
        /// </summary>
        public ProcessorFactory(ISldWorks swApp)
        {
            _swApp = swApp;
            _generic = new GenericPartProcessor();
            _sheet = new SheetMetalPartProcessor(swApp);
            _tube = new TubePartProcessor(swApp);
            _machined = null; // Not implemented yet
        }

        /// <summary>
        /// Full constructor for testing or custom processor injection.
        /// </summary>
        public ProcessorFactory(IPartProcessor generic, IPartProcessor sheet = null, IPartProcessor tube = null, IPartProcessor machined = null)
        {
            _swApp = null;
            _generic = generic ?? new GenericPartProcessor();
            _sheet = sheet;
            _tube = tube ?? new TubePartProcessor();
            _machined = machined;
        }

        public IPartProcessor Get(ProcessorType type)
        {
            switch (type)
            {
                case ProcessorType.SheetMetal: return _sheet ?? _generic;
                case ProcessorType.Tube: return _tube ?? _generic;
                case ProcessorType.Machined: return _machined ?? _generic;
                default: return _generic;
            }
        }

        /// <summary>
        /// Quick detection based on existing features (original behavior).
        /// Fast path for parts that already have sheet metal or tube features.
        /// </summary>
        public IPartProcessor DetectFor(IModelDoc2 model)
        {
            if (_sheet != null && _sheet.CanProcess(model)) return _sheet;
            if (_tube != null && _tube.CanProcess(model)) return _tube;
            if (_machined != null && _machined.CanProcess(model)) return _machined;
            return _generic;
        }

    }
}
