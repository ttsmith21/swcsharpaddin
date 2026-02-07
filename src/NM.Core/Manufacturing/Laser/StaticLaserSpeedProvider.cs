using NM.Core.Config;
using NM.Core.Config.Tables;

namespace NM.Core.Manufacturing.Laser
{
    /// <summary>
    /// Laser speed/pierce data provider backed by nm-tables.json.
    /// Falls back to compiled defaults in NmTables if JSON is not loaded.
    /// Matching algorithm: VBA parity — find thinnest entry where thickness >= (partThickness - tolerance).
    /// </summary>
    public sealed class StaticLaserSpeedProvider : ILaserSpeedProvider
    {
        public LaserSpeed GetSpeed(double thicknessIn, string materialCode)
        {
            var (feedIpm, pierceSec) = NmTablesProvider.GetLaserSpeed(
                NmConfigProvider.Tables, thicknessIn, materialCode);

            return new LaserSpeed
            {
                FeedRateIpm = feedIpm,
                PierceSeconds = pierceSec
            };
        }
    }
}
