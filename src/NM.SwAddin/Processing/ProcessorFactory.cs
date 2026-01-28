using NM.Core;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin.Processing
{
    public interface IProcessorFactory
    {
        IPartProcessor Get(ProcessorType type);
        IPartProcessor DetectFor(IModelDoc2 model);

        /// <summary>
        /// Detects and routes to appropriate processor using full classification pipeline.
        /// Use this for imported solids that don't have existing features.
        /// </summary>
        IPartProcessor DetectAndClassify(IModelDoc2 model, ModelInfo info, ProcessingOptions options, out ClassificationResult classification);
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

        /// <summary>
        /// Full classification using trial-and-validate pipeline.
        /// Use for imported solids that need classification.
        /// </summary>
        /// <param name="model">The model to classify.</param>
        /// <param name="info">ModelInfo for property updates.</param>
        /// <param name="options">Processing options.</param>
        /// <param name="classification">Output: The classification result with extracted metrics.</param>
        /// <returns>Appropriate processor for the classification, or generic if classification failed.</returns>
        public IPartProcessor DetectAndClassify(IModelDoc2 model, ModelInfo info, ProcessingOptions options, out ClassificationResult classification)
        {
            classification = null;

            // Fast path: if already has features, use quick detection
            if (_sheet != null && _sheet.CanProcess(model))
            {
                classification = new ClassificationResult(PartClassification.SheetMetal, "Existing sheet metal features");
                return _sheet;
            }
            if (_tube != null && _tube.CanProcess(model))
            {
                classification = new ClassificationResult(PartClassification.Tube, "Existing tube detection");
                return _tube;
            }

            // No existing features - use full classification pipeline
            if (_swApp == null)
            {
                ErrorHandler.HandleError("ProcessorFactory.DetectAndClassify", "ISldWorks not available for classification pipeline", null, ErrorHandler.LogLevel.Warning);
                classification = new ClassificationResult(PartClassification.Other, "Classification unavailable");
                return _generic;
            }

            var pipeline = new ClassificationPipeline(_swApp);
            classification = pipeline.Classify(model, info, options);

            switch (classification.Classification)
            {
                case PartClassification.SheetMetal:
                    // Sheet metal conversion already happened in pipeline
                    // Return sheet processor for any additional processing
                    return _sheet ?? _generic;

                case PartClassification.Tube:
                    return _tube ?? _generic;

                case PartClassification.Other:
                case PartClassification.Failed:
                default:
                    return _generic;
            }
        }
    }
}
