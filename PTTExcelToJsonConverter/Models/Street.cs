using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PTTExcelToJsonConverter.Models
{
    public class Street
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string PK { get; set; }
        public int CityId { get; set; }
        public int DistrictId { get; set; }
        public City City { get; set; } = new City();
        public District District { get; set; } = new District();
    }
}
