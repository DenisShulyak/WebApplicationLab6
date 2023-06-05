using System;
using WebApplicationLab6.Objects.Abstracts;

namespace WebApplicationLab6.Objects
{
    public class ContractCity : Entity
    {
        public Guid ContractId { get; set; }
        
        public Contract Contract { get; set; }
        
        public Guid CityId { get; set; }
        
        public City City { get; set; }
        
        public decimal Price { get; set; }
    }
}