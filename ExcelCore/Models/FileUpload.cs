namespace ExcelCore.Models
{
    public class FileUpload
    {
     
        public List<IFormFile> XlsFiles { get; set; }
        
        public Info InfoModel { get; set; }

        public YInfo YoutubeModel { get; set; }

        public DInfo DjezzyModel { get; set; }

        public OInfo OoredooModel { get; set; }

        public ZInfo ZainIQModel { get; set; }


        public RbtNames rbtNames { get; set; }

        public DateTime rbtDate { get; set; }

        public FileUpload()
        {
            
            InfoModel = new Info();
            OoredooModel = new OInfo();
            YoutubeModel = new YInfo();
            DjezzyModel = new DInfo();
            ZainIQModel = new ZInfo();
            rbtNames = new RbtNames();
            rbtDate = DateTime.Today;


        }
    }
}
