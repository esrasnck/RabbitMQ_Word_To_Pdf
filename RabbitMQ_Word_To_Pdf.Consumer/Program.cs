using Newtonsoft.Json;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using Spire.Doc;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace RabbitMQ_Word_To_Pdf.Consumer
{

    // bu derse email kısmını halledecez. o yüzden email için bir metot oluşturcaz.
    class Program
    {
        static void Main(string[] args)
        {
            bool result = false;

            var factory = new ConnectionFactory();
            factory.Uri = new Uri("amqps://evrdybgg:TQLBLVPPb7_OR1qBR3hajFa7ncOrB_HM@fish.rmq.cloudamqp.com/evrdybgg");

            using (var connection = factory.CreateConnection())  // 1) bağlantı oluşturdum.
            {
                using (var channel = connection.CreateModel())  // 2) kanal oluşturudum
                {
                    channel.ExchangeDeclare("convert-exchange", ExchangeType.Direct, true, false, null);

                    channel.QueueBind(queue:"File",exchange: "convert-exchange", "WordToPdf", null); // routing key vermeyince, null refence ex. alıyor

                    channel.BasicQos(0,1,false); // mesajların eşit bir biçimde dağılması için

                    var consumer = new EventingBasicConsumer(channel);

                    channel.BasicConsume("File",false,consumer); // false => eposta gönderemezsem, o kuyruktan silinmesin. başka bir instance varsa ona atsın.

                    consumer.Received += (model, ea) =>
                    {
                        // dönüştürme başarılı olmazsa
                        try
                        {
                            Console.WriteLine("Kuyruktan bir mesaj alındı ve işleniyor.");

                            // dönüştürme işlemi sağlamak için, spire.doc'ın kütüphanesini kullancam.

                            Document document = new Document();
                            //ea üzerinden datayı alacam.
                            string messsage = Encoding.UTF8.GetString(ea.Body.ToArray()); // döncez ?? Array deyince yedi?

                            MessageWordToPdf messageWordToPdf = JsonConvert.DeserializeObject<MessageWordToPdf>(messsage); // gelen mesajı deserialize ediyoruz.

                            document.LoadFromStream(new MemoryStream(messageWordToPdf.WordByte),FileFormat.Docx2013); // ben stream üzerinden gittiğim için bunun üzerinden yükleyeceğim.

                            using (MemoryStream ms = new MemoryStream())
                            {
                                document.SaveToStream(ms, FileFormat.PDF); // stream olan pdf formatını buraya kaydediyor

                                result = EmailSender(messageWordToPdf.Email,ms,messageWordToPdf.FileName); // result işlemi başarısızsa ne olacak? başarılıysa bana true dönecek

                            };




                        }
                        catch (Exception ex)
                        {

                            Console.WriteLine("Hata meydana geldi: " + ex.Message);
                        }

                        if(result) // mesajın gönderilme ile ilgili durumu.
                        {
                            Console.WriteLine("kuyruktan mesaj başarıyla işlendi...");
                            // simdi de bunu rabbitMq ya göndercem
                            channel.BasicAck(ea.DeliveryTag, false); // bu doğru ise, artık bu mesajı kuyruktan silebilirsin diyoruz.

                        }


                    };

                    Console.WriteLine("Cıkmak için tıklayınız.");
                    Console.ReadLine();
                }


            }


        }

        // memoryStream içinde bir pdf dosyamız var
        public static bool EmailSender(string email, MemoryStream memoryStream, string fileName)
        {
            try
            {
                memoryStream.Position = 0; // 0'dan itibaren memory stream dosyasını ilgili epostaya iliştir diyorum. bu kodu kullanmazsak, pdf dosyası 0 byte olarak gider. cünkü position'ı oluşturamamıştır ve doğal olarak da email'i bağlayamamıştır. bu kodun anlamı=> "ilk satırdan itibaren okumaya başlaması"

                // bağlayacağım noktanın tipinin pdf olduğunu belirtiyoruz.

                System.Net.Mime.ContentType content_type = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Application.Pdf);

                Attachment attach = new Attachment(contentStream: memoryStream, contentType: content_type);
                // dosya ismini beliriycem (attach'in)

                attach.ContentDisposition.FileName = $"{fileName}.pdf";

                // smpt protokolü ayarlarını yapacağız



                SmtpClient smtpClient = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential("bayms3132@gmail.com", "password")
                    
                    //todo:testten önce şifreni yaz.
                };

                MailMessage mailMessage = new MailMessage();

                mailMessage.From = new MailAddress("bayms3132@gmail.com");  // fake mail adresim

                mailMessage.To.Add(email); // gönderceğim mail

                mailMessage.Subject = "Pdf Dosyası oluşturuldu...";

                mailMessage.Body = "Pdf dosyanız ektedir";

                mailMessage.IsBodyHtml = true;

                // artık mail mesajı attach edecem. bağlıycam

                mailMessage.Attachments.Add(attach);

                smtpClient.Send(mailMessage);

                Console.WriteLine($"Sonuc: {email} adresine gönderilmiştir. ");

                // memory'i kapatıp dispose ediyoruz.

                memoryStream.Close();
                memoryStream.Dispose();

                return true;
            }
            catch (Exception ex)
            {

                Console.WriteLine($"Mail gönderim sırasında bir hata meydana geldi : {ex.InnerException}");
                return false;
            }




        }
    }
}


