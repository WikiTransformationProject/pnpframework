using Newtonsoft.Json;
using System.Xml.Serialization;

namespace PnP.Framework.Provisioning.Model
{
    // written by LLM, 2026-11-15
    public abstract partial class BaseClientSidePage
    {
        /// In-memory receipt set after the page writer completes successfully.
        /// The caller clears it before each apply. It is not part of a stored template.
        /// Page templates must retain the supplied page model to return this receipt.
        [JsonIgnore]
        [XmlIgnore]
        public bool WikiTraccsPageWriteSucceeded { get; set; }
    }
}
