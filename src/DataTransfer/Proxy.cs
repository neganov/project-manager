using Microsoft.ServiceBus;
using System.ServiceModel;
using Dispatcher;

namespace DataTransfer
{
    public static class Proxy
    {
        private static ChannelFactory<IEventChannel> _ChannelFactory;

        public static void PushMessage(ModificationDto dto)
        {
            if(_ChannelFactory == null)
            {
                _ChannelFactory = new ChannelFactory<IEventChannel>(
                    new NetTcpRelayBinding(),
                    "sb://projectmanagementprototype.servicebus.windows.net/events");
                _ChannelFactory.Endpoint.EndpointBehaviors.Add(
                    new TransportClientEndpointBehavior
                    {
                        TokenProvider = TokenProvider.CreateSharedAccessSignatureTokenProvider(
                            "RootManageSharedAccessKey",
                            "v0M0rakA8q03aSetYC00EJI8s1bVzrI5InMu6VyKIZI=")
                    });
            }

            using (IEventChannel channel = _ChannelFactory.CreateChannel())
            {
                var modification = new ProjectModification
                {
                    ProjectId = dto.ProjectId,
                    OriginalMilestone = dto.OriginalMilestone,
                    NewMilestone = dto.NewMilestone
                };

                channel.OnProjectChanged(modification);
            }
        }
    }
}
