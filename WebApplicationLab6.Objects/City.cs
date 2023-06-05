using System.Collections.Generic;
using WebApplicationLab6.Objects.Abstracts;

namespace WebApplicationLab6.Objects
{
    public class City : Entity
    {
        public string Name { get; set; }
        
        public ICollection<ContractCity> ContractCities { get; set; }
    }
}