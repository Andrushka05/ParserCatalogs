using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParserCatalog
{
    public class Site
    {
        public Site()
        {
            Catalog = false;
            Categories=new List<Category>();
        }
        public string Name { get; set; }
        public List<Category> Categories { get; set; }
        public bool Catalog { get; set; }
    }
}
