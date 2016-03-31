using System.Collections.Generic;
using Microsoft.AspNet.Mvc;
using DataTransfer;

namespace Provisioner.Controllers
{
    [Route("api/[controller]")]
    public class ModificationsController : Controller
    {

        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "modification1", "modification2" };
        }

        [HttpGet("{id}")]
        public string Get(string id)
        {
            return "a modification";
        }

        [HttpPost]
        public void Post([FromBody]ModificationDto dto)
        {
            Proxy.PushMessage(dto);
        }
    }
}
