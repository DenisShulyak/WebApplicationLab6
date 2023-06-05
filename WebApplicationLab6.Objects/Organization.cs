using System;
using WebApplicationLab6.Objects.Abstracts;

namespace WebApplicationLab6.Objects
{
    public class Organization : Entity
    {
        public string Name { get; set; }
        
        public string Inn { get; set; }
        
        public string Kpp { get; set; }
        
        public string Address { get; set; }
        
        public Guid CityId { get; set; }
        
        public City City { get; set; }
        
        public Guid OrganizationTypeId { get; set; }
        
        public OrganizationType OrganizationType { get; set; }
    }
}