using NM.Core.Config;
using NM.Core.Config.Tables;

namespace NM.Core.Tubes
{
    /// <summary>
    /// Provides cutting parameters for tube processing, backed by nm-tables.json.
    /// Units: inches and seconds. Feed rates are inches/min.
    /// </summary>
    public sealed class TubeCuttingParameterService
    {
        public sealed class CutParams
        {
            public double KerfIn { get; set; }
            public double PierceTimeSec { get; set; }
            public double CutSpeedInPerMin { get; set; }
        }

        /// <summary>
        /// Returns parameters based on material category and wall thickness.
        /// materialCategory: "CarbonSteel", "StainlessSteel", "Aluminum" (case-insensitive).
        /// </summary>
        public CutParams Get(string materialCategory, double wallIn)
        {
            var (cutSpeed, pierceSec, kerfIn) = NmTablesProvider.GetTubeCuttingParams(
                NmConfigProvider.Tables, materialCategory, wallIn);

            return new CutParams
            {
                KerfIn = kerfIn,
                PierceTimeSec = pierceSec,
                CutSpeedInPerMin = cutSpeed
            };
        }
    }
}
