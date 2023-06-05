using System;
using WebApplicationLab6.Objects.Abstracts;

namespace WebApplicationLab6.Objects
{
    public class User : Entity
    {
        public string Login { get; set; }
        
        public string Password { get; set; }
        
        public string FirstName { get; set; }
        
        public string LastName { get; set; }
        
        public string Patronymic { get; set; }
        
        public string Snils { get; set; }
        
        public DateTime BirthDate { get; set; }
        
        public Guid RoleId { get; set; }
        public Role Role { get; set; }
        
        public Guid OrganizationId { get; set; }
        
        public Organization Organization { get; set; }
    }
}