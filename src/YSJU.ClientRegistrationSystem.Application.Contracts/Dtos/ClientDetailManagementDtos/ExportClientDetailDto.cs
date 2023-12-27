using System;
using System.Collections.Generic;
using System.Text;

namespace YSJU.ClientRegistrationSystem.Dtos.ClientDetailManagementDtos
{
    public class ExportClientDetailDto
    {
        public string Name { get; set; }
        public byte[] content { get; set; }
    }
}
