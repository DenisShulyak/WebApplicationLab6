using System;
using System.ComponentModel.DataAnnotations;

namespace WebApplicationLab6.Objects.Abstracts
{
    public class Entity
    {
        [Key]
        public Guid Id { get; set; }
    }
}