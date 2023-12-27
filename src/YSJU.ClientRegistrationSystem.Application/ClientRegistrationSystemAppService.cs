using System;
using System.Collections.Generic;
using System.Text;
using YSJU.ClientRegistrationSystem.Localization;
using Volo.Abp.Application.Services;

namespace YSJU.ClientRegistrationSystem;

/* Inherit your application services from this class.
 */
public abstract class ClientRegistrationSystemAppService : ApplicationService
{
    protected ClientRegistrationSystemAppService()
    {
        LocalizationResource = typeof(ClientRegistrationSystemResource);
    }
}
