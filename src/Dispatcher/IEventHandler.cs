using System.ServiceModel;

namespace Dispatcher
{
    [ServiceContract]
    public interface IEventHandler
    {   
        [OperationContract]
        void OnProjectChanged(ProjectModification modification);
    }
}
