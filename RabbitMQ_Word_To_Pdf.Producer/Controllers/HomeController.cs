using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using RabbitMQ.Client;
using RabbitMQ_Word_To_Pdf.Producer.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RabbitMQ_Word_To_Pdf.Producer.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult WordToPdfPage()
        {
            return View();
        }

        [HttpPost]
        public IActionResult WordToPdfPage(WordToPdf wordToPdf)
        {
            // önce connect olacağız
            var factory = new ConnectionFactory();
            factory.Uri = new Uri(_configuration["ConnectionStrings:RabbitMQ"]);

            using(var connection = factory.CreateConnection())  // 1) bağlantı oluşturdum.
            { 
                using(var channel = connection.CreateModel())  // 2) kanal oluşturudum
                {
                    channel.ExchangeDeclare("convert-exchange",ExchangeType.Direct,true,false,null);     // 3) exchange oluşturdum(durable: true, (fiziksel olarak kaydedilsin. restart olduğunda silinmesin.) aoutdelete false. arguman ise null)

                    channel.QueueDeclare(queue: "File", durable: true, exclusive: false, autoDelete: false);// 4) kuyruk oluşturdum. => göndermiş olduğum dosyaların kuyrukta tutulmasını istiyorum. (exclusive: false=> birden fazla bağlantı kuyruğu kulllanabilsin

                    channel.QueueBind("File", "convert-exchange","WordToPdf"); // 5) kuyruk ile exchange i bind edecem (mesajımın kaybolmasını istemiyorum.) WordToPdf = routingKey

                    /*benim artık modele ihtiyacım var. complex tip / sınıf göndercem. ben bu sınıfta, word dosyası, email adresi ve dosyanın adını göndercem.
                      eğer kullancı benim sistemimdeki word dosyasını kullandıysa, ben kullanıcıya email atmam gerek. işte oradaki dosya ismini de almam gerek
                      ismi neyse o isimde oluşmuş bir pdf dosyası göndermem gerek. bu yüzden bir class daha oluşturuyorum... o yüzden message classı=> şair burada ne dedi anlamadım. ? */

                    // wordToPdf=> kullanıcının gönderdiği pdf dosyasının içindeki class. MessageWordToPdf => kullanıcıya mesaj gönderdiğimiz classs

                    MessageWordToPdf messageWordToPdf = new MessageWordToPdf();

                    using (MemoryStream memoryStream = new MemoryStream())// bu word dosyasını memoryde tutacam
                    {
                        wordToPdf.WordFile.CopyTo(memoryStream); // hafızaya kopyaladım
                        messageWordToPdf.WordByte = memoryStream.ToArray(); // metoduma gelen wordToPdf dosyasını memory kopyalıyoruz. sonra memory nesnesini dolayısıyla onu array'e çevirip, mesajı göndereceğim sınıfa atıyoruz.

                    }

                    messageWordToPdf.Email = wordToPdf.Email;
                    messageWordToPdf.FileName = Path.GetFileNameWithoutExtension(wordToPdf.WordFile.FileName); // file name'i alacaz (uzantısı olmadan dosya adını alıyoruz.) -- wordToPdf.WordFile.FileName => kullanıcıdan aldığım dosyanın file adi.

                    // complex tipleri serilaze etmem gerek

                    string serializeMessage = JsonConvert.SerializeObject(messageWordToPdf);

                    // sonra bunu byte a çevircem

                    byte[] byteMessage = Encoding.UTF8.GetBytes(serializeMessage);

                    // mesajin sabit bicimde gitmesi icin ilk olarak, kuyruğu sabitlemem gerek. durabile i true yaparak bunu yaparım. ikincisi de property'inin persistance'ını true ya set edecem

                    var properties = channel.CreateBasicProperties();

                    properties.Persistent = true;

                    channel.BasicPublish(exchange:"convert-exchange",routingKey: "WordToPdf",basicProperties:properties,body:byteMessage);  // burada da artık gönderiyoruz.

                    ViewBag.result = "Word dosyanız, pdf dosyasına dönüştürüldükten sonra, size email olarak gönderilecektir.";

                    return View();

                }
            }
          

            return View();
        }


        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
