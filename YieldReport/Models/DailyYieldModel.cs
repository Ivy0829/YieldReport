using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace YieldReport.Models
{
    public class DailyYieldModel
    {

        public string Guid { get; set; }
        public string YearCode { get; set; }
        public string Plant { get; set; }
        public string SubLotNo { get; set; }
        public string LotNo { get; set; }
        public string StageCode { get; set; }
        public string Cust2Code { get; set; }
        public string Cust3Code { get; set; }
        public string PkgCode { get; set; }
        public string Device { get; set; }
        public DateTime TrackInTime { get; set; }
        public int TrackInQty { get; set; }
        public DateTime TrackOutTime { get; set; }
        public int TrackOutQty { get; set; }
        public int SumDefectQty { get; set; }
        public string RunType { get; set; }
        public string Yield { get; set; }



















    }
}
