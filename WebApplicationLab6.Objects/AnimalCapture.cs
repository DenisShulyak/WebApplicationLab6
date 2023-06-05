using System;
using WebApplicationLab6.Objects.Abstracts;

namespace WebApplicationLab6.Objects
{
    public class AnimalCapture : Entity
    {
        public string Property { get; set; }
        
        public string Animal { get; set; }
        
        public string Ears { get; set; }
        
        public string Tail { get; set; }
        
        public string Features { get; set; }
        
        public string IdentityMark { get; set; }
        
        public string EChipNum { get; set; }
        
        public Guid CaptureActId { get; set; }
        
        public CaptureAct CaptureAct { get; set; }
        
        public Guid CityId { get; set; }
        
        public City City { get; set; }
    }
}