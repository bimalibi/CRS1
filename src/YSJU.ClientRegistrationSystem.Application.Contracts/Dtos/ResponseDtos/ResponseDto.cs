using System;
using System.Collections.Generic;
using System.Text;

namespace YSJU.ClientRegistrationSystem.Dtos.ResponseDtos
{
    public class ResponseDto<T>
    {
        public bool Success { get; set; }
        public int Code { get; set; }
        public string Message { get; set; }
        public T? Data { get; set; }
    }
}
