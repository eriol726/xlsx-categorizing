using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SortXLSX
{
    public class Category
    {
        public CategoryItem Hyra = new CategoryItem { title = "Hyra", sum = 0.0 };
        public CategoryItem Bredband = new CategoryItem { title = "Bredband/Modil", sum = 0 };
        public CategoryItem Manadsutgifter = new CategoryItem { title = "Månadsutgifter", sum = 0 };
        public CategoryItem Swish = new CategoryItem { title = "Swish", sum = 0 };
        public CategoryItem Mat = new CategoryItem { title = "Mat", sum = 0 };
        public CategoryItem Sprit = new CategoryItem { title = "Sprit", sum = 0 };
        public CategoryItem Resor = new CategoryItem { title = "Resor", sum = 0 };
        public CategoryItem Other = new CategoryItem { title = "Övrigt", sum = 0 };
        public CategoryItem Snus = new CategoryItem { title = "Snus", sum = 0 };
    }
}


