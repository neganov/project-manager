using System.Runtime.Serialization;

namespace Dispatcher
{
    [DataContract]
    public class ProjectModification
    {
        [DataMember]
        public string ProjectId { get; set; }
        [DataMember]
        public string OriginalMilestone { get; set; }
        [DataMember]
        public string NewMilestone { get; set; }
    }
}
