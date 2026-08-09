using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System;
using System.Threading.Tasks;
using Microsoft.ServiceHealthHub.Core;
using Microsoft.AspNetCore.Authorization;

namespace Service_Health_Dashboard_v2.Controllers
{
    public class MSImage
    {
        private int _id;
        private string _name;
        private string _type;
        private string _format;
        private string _content;

        public int Id => _id;
        public string Name => _name;
        public string Type => _type;
        public string Format => _format;
        public string Content => _content;

        public static List<MSImage> GetImages()
        {
            return GetImages("");
        }
            
        public static List<MSImage> GetImages(string type)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            List<MSPropertyBag> images = db.GetImages(type);
            List<MSImage> result = new List<MSImage>();

            foreach (MSPropertyBag image in images)
            {
                MSImage imageObject = new MSImage
                {
                    _id = image["Id"] is DBNull ? -1 : (int)image["Id"],
                    _name = image["Name"] is DBNull ? string.Empty : (string)image["Name"],
                    _type = image["Type"] is DBNull ? string.Empty : (string)image["Type"],
                    _format = image["Format"] is DBNull ? string.Empty : (string)image["Format"],
                    _content = image["Content"] is DBNull ? string.Empty : (string)image["Content"]
                };

                result.Add(imageObject);
            }

            return result;
        }

        public static MSImage GetImage(string name, string type)
        {
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            List<MSPropertyBag> images = db.GetImage(name, type);

            if (images == null || images.Count <= 0)
                return null;
            else
            {
                MSPropertyBag image = images[0];
                MSImage result = new MSImage
                {
                    _id = image["Id"] is DBNull ? -1 : (int)image["Id"],
                    _name = image["Name"] is DBNull ? string.Empty : (string)image["Name"],
                    _type = image["Type"] is DBNull ? string.Empty : (string)image["Type"],
                    _format = image["Format"] is DBNull ? string.Empty : (string)image["Format"],
                    _content = image["Content"] is DBNull ? string.Empty : (string)image["Content"]
                };

                return result;
            }
        }
    }

    [Route("api/[controller]")]
    [Authorize(Roles = "ServiceHealthReader,Communication.Write.All,Public,LicenseReader,Admin")]
    [ApiController]
    public class ImageStore : ControllerBase
    {
        private readonly ILogger<PublicReport> _logger;

        public ImageStore(ILogger<PublicReport> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public async Task<List<MSImage>> Get(string name, string type)
        {
            if (string.IsNullOrWhiteSpace(type))
                return MSImage.GetImages();
            else
                return MSImage.GetImages(type);
        }
    }
}
