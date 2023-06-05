using System;
using WebApplicationLab6.Objects.Abstracts;

namespace WebApplicationLab6.Objects
{
    public class CaptureAct : Entity
    {
        public int CountDogs { get; set; }
        
        public int CountCats { get; set; }
        
        public int CountAnimals { get; set; }
        
        public DateTime CaptureDate { get; set; }
        
        public string CapturePurpose { get; set; }
        
        public Guid OrganizationId { get; set; }
        
        public Organization Organization { get; set; }
        
        public Guid ContractId { get; set; }
        
        public Contract Contract { get; set; }
        
        public Guid ClaimId { get; set; }
        
        public Claim Claim { get; set; }
    }
}