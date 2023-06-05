using System;
using System.Collections.Generic;
using WebApplicationLab6.Objects.Abstracts;

namespace WebApplicationLab6.Objects
{
    public class Contract : Entity
    {
        public DateTime BeginDate { get; set; }
        
        public DateTime EndDate { get; set; }
        
        public  Organization Customer { get; set; }
        
        public  Guid CustomerId { get; set; }
        
        public  Organization Executor { get; set; }
        
        public  Guid ExecutorId { get; set; }
        
        public ICollection<ContractCity> ContractCities { get; set; }
    }
}