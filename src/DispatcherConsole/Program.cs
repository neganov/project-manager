using System;
using System.ServiceModel;

namespace DispatcherConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var host = new ServiceHost(typeof(Dispatcher.EventHandler));
            host.Open();
            Console.WriteLine("\r\nListening to incoming events. Press ENTER to exit...\r\n");
            Console.ReadLine();
            host.Close();
        }
    }
}
