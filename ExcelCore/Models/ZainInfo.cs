namespace ExcelCore.Models
{
    public class ZInfo
    {
        internal readonly IEnumerable<object> InfoLists;

        public string date { get; set; }

        public string RBTID { get; set; }
        public string RBTName { get; set; }
        public string Artist { get; set; }
        public string NumberofDownloads { get; set; }
        public string NumberofPDownloads { get; set; }
        public string TotalRevenue { get; set; }

        public string Category { get; set; }

        public List<ZInfo> InfoList { get; set; }
        public ZInfo()
        {
            InfoList = new List<ZInfo>();
        }


    }
}