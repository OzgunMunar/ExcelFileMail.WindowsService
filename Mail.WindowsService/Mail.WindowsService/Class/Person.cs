using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mail.WindowsService.DTOClasses
{
    public class Person
    {

        public string PersonName { get; set; }
        public string Address { get; set; }
        public string Nation { get; set; }
        public string CarPlate { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime FinishDate { get; set; }
        public string Notes { get; set; }

    }
}
