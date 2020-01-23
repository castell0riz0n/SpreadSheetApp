using System.Collections.Generic;

namespace ExcelApp.Classes
{
    public class ResourcesBaseDto
    {
        public List<ResourcesDto> Sheet1 { get; set; }
        public List<ResourcesDto> Sheet2 { get; set; }
    }



    public class ResourcesDto
    {
        public string Name { get; set; }
        public string Value { get; set; }
        public string Comment { get; set; }
    }
}