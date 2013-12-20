using System.Collections.Generic;
using System.Drawing;

namespace ParserCatalog
{
    public class Product
    {
        public string Color { get; set; }
        public string Article { get; set; }
        public string Size { get; set; }
        public string CategoryPath { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string Price { get; set; }
        public string Url { get; set; }
        public string Photo { get; set; }
        public List<string> Photos { get; set; }
     }
}
