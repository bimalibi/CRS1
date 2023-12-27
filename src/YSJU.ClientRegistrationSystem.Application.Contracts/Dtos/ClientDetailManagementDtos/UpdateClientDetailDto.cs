using System;
using System.Collections.Generic;
using System.Text;

namespace YSJU.ClientRegistrationSystem.Dtos.ClientDetailDtos
{
    public class UpdateClientDetailDto
    {
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string LastName { get; set; }
        public string Address { get; set; }
        public string PhoneNumber { get; set; }
        public List<Guid> ProductCategoryId { get; set; }
    }
}
