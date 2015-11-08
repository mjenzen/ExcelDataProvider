using Excel.Data.Sample.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel.Data.Sample
{
    public class SampleContext : ExcelDataProvider
    {
        public IEnumerable<Brand> Brands { get { return GetObject<Brand>(); } }
        public IEnumerable<Category> Categories { get { return GetObject<Category>(); } }

        public SampleContext(string filepath) : base(filepath)
        {
        }
    }
}
