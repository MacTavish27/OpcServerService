using System;
using System.Web.Http;
using static opc_bridge.Services;

namespace opc_bridge
{
    [RoutePrefix("api/opctags")]
    public class OpcTagsController : ApiController
    {
        private OpcServerService Service => OpcServerService.Instance;

        [HttpGet]
        [Route("{tagId}")]
        public IHttpActionResult GetTag(string tagId)
        {
            try
            {
                var result = Service.ReadTag(tagId);

                if (result == null)
                    return NotFound();


                return Ok(new
                {
                    tagId = result.ItemName,
                    value = result.Value,
                    quality = result.Quality.ToString(),
                    timestamp = result.Timestamp
                });


            }

            catch (Exception ex)
            {
                HttpLogger.Log("[ERROR] Error in GET request: " + ex.Message);
                return InternalServerError(ex);
            }

        }

        [HttpPost]
        [Route("subscribe")]
        public IHttpActionResult Subscribe([FromBody] string[] tagIds)
        {

            Service.SubscribeTags(tagIds);

            return Ok($"Subscribed {tagIds.Length} tags");

        }


        [HttpPost]
        [Route("")]
        public IHttpActionResult WriteTag([FromBody] TagWriteModel model)
        {
            try
            {
                Service.WriteTag(model.TagId, model.Value);
                return Ok("Tag written successfully");
            }
            catch (Exception ex)
            {
                HttpLogger.Log("[ERROR] Error in POST request: " + ex.Message);
                return InternalServerError(ex);
            }
        }

        public class TagWriteModel
        {
            public string TagId { get; set; }
            public object Value { get; set; }
        }
    }
}
