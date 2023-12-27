using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using Volo.Abp.Application.Dtos;

namespace YSJU.ClientRegistrationSystem.Dtos.ClientDetailDtos
{
    public class ClientDetailResponseDto : EntityDto<Guid>
    {
        public int ClientId { get; set; }
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string LastName { get; set; }
        public string Address { get; set; }
        public string PhoneNumber { get; set; }
        public string Email { get; set; }
        public Guid ProductCategoryId { get; set; }
        public string ProductCategoryName { get; set; }
        public DateTime CreationTime { get; set; }
    }
}
