using System.ServiceModel;

namespace Dispatcher
{
    public interface IEventChannel : IEventHandler, IClientChannel
    {
    }
}
