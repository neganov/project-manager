using System;

namespace Dispatcher
{
    public class EventHandler : IEventHandler
    {
        public void OnProjectChanged(ProjectModification modification)
        {
            if(modification == null)
            {
                Console.WriteLine("Received a 'NULL' modification.");
            }
            else
            {
                Console.WriteLine($"Received modification for project id={modification.ProjectId}:");
                Console.WriteLine($"\tOriginal milestone: {modification.OriginalMilestone}");
                Console.WriteLine($"\tNew milestone: {modification.NewMilestone}");
            }
        }
    }
}
