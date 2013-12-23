using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParserCatalog
{
    public class ShopBig:Shop
    {
        public ShopBig()
        {
            CatalogList=new List<Category>();
        }
        public List<Category> CatalogList { get; set; } 
        public string XPath { get; set; }
        public string Host { get; set; }
    }
}
