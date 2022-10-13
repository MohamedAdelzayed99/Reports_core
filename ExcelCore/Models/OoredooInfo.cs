namespace ExcelCore.Models
{
    public class OInfo
    {
        internal readonly IEnumerable<object> InfoLists;

        public string date { get; set; }

        public string ContentName { get; set; }
        public string Artist { get; set; }
        public string MediaId { get; set; }
        public string Contentprovider { get; set; }
        public string Trackpkey { get; set; }
        public string CatName { get; set; }
        public string ProviderName { get; set; }
        public string PricePointValue { get; set; }
       

        public List<OInfo> InfoList { get; set; }
        public OInfo()
        {
            InfoList = new List<OInfo>();
        }

       
    }
}
