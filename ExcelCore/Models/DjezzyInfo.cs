﻿namespace ExcelCore.Models
{
    public class DInfo
    {
        public string date { get; set; }

        public string RankID { get; set; }
        public string ResourceName { get; set; }
        public string ResourceCode { get; set; }
        public string ISRC { get; set; }
        public string SPName { get; set; }
        public string SPCode { get; set; }
        public string Artist { get; set; }
        public string ChargedPrice { get; set; }
        public string DownloadTimes { get; set; }
        public string CopyTimes { get; set; }
        public string PresentTimes { get; set; }
        public string RenewalTimes { get; set; }
        public string TotalTimes { get; set; }


        public List<DInfo> InfoLists { get; set; }
            public DInfo()
            {
                InfoLists = new List<DInfo>();
            }
        }
    }

