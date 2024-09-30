
using System.Net;
using System.Net.Mail;
using System.Net.Mime;

public class Program
{
    static void Main(String[] args)
    {

        string host = "outlook.office365.com";
        int port = 143;
        string username = "P00993439@asianpaints.com";
        string password = "GO$g18LGr";
        string from = "P00993439@asianpaints.com";
        string to = "nithish.k@orai-robotics.com";
        //APGWABOT@asianpaints.com

        //string host = "smtp.office365.com";
        //int port = 587;
        //string username = "support@orai-robotics.com";
        //string password = "Orai@1234#";
        //string from = "support@orai-robotics.com";
        //string to = "nithish.k@orai-robotics.com";

        string localIp = Dns.GetHostAddresses(Dns.GetHostName())[0].ToString();
        string strHostName = Dns.GetHostName();
        string ji= $"\n\nLocal IP: {localIp}\nDestination IP: {host}";
        IPHostEntry ipEntry = Dns.GetHostEntry(strHostName);
        IPAddress[] addr = ipEntry.AddressList;
        for (int i = 0; i < addr.Length; i++)
        {
            Console.WriteLine("IP Address {0}: {1} ", i, addr[i].ToString());
        }

        using (var smtpClient = new SmtpClient(host))
        {
            smtpClient.Port = port;
            smtpClient.Credentials = new NetworkCredential(username, password);
            smtpClient.EnableSsl = true;

            HttpClient client = new HttpClient();
            HttpResponseMessage http = client.GetAsync("https://fpi.branding-element.com/prod/97326/WA_PUBLIC_ATTACHMENT/8ee7376f_58d8_4258_8f8b_51d8b73f50ec.ogg-6zeNq.ogg").Result;
            byte[] bytes = http.Content.ReadAsByteArrayAsync().Result;
            MemoryStream mediaStream = new MemoryStream(bytes);

             //bytes
            Attachment attachment = new Attachment(mediaStream, "test.ogg", "audio/ogg");


            using (var mailMessage = new MailMessage(from, to))
            {
                mailMessage.Subject = "Test Subject";
                mailMessage.Body = "Test Body";
                mailMessage.Attachments.Add(attachment);
                smtpClient.Send(mailMessage);
                Console.WriteLine("Email sent successfully!");
            }
        }

    }
}