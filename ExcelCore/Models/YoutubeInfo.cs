
   namespace ExcelCore.Models
    {
        public class YInfo
        {
            public string date { get; set; }

            public string Asset { get; set; }
            public string Yourer { get; set; }
            public string YourYPr { get; set; }
            public string Yourtr { get; set; }
           

            public List<YInfo> InfoLists { get; set; }
            public YInfo()
            {
                InfoLists = new List<YInfo>();
            }
        }
    }
