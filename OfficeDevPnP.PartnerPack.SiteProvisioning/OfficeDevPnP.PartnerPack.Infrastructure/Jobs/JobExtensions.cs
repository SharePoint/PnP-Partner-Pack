using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Jobs
{
    /// <summary>
    /// Provides extension methods for Provisioning Jobs
    /// </summary>
    public static class JobExtensions
    {
        public static Stream ToJsonStream(this ProvisioningJob job)
        {
            String jsonString = JsonConvert.SerializeObject(job);
            Byte[] jsonBytes = System.Text.Encoding.Unicode.GetBytes(jsonString);
            MemoryStream jsonStream = new MemoryStream(jsonBytes);
            jsonStream.Position = 0;

            return (jsonStream);
        }
    }
}
