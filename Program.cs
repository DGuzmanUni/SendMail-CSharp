using System;
using MailKit.Net.Smtp;
using MailKit;
using MailKit.Security;
using MimeKit;
using MimeKit.Text;



namespace SendMailEx{
    class Program
    {
        static void Main(string[ ] args)
        {
            // Credenciales del emisor
            string emailFrom = "email@From.com";
            string passwordFrom = "1234";

            // Email Receptor
            string emailTo = "email@receptor.com";
            
            // Cuerpo del Mensaje 
            string emailBody = "Body Message";

            MimeMessage message = new MimeMessage();

            // ->Crear el mensaje
            // emisor
            message.From.Add(MailboxAddress.Parse(emailFrom));
            // receptor
            message.To.Add(MailboxAddress.Parse(emailTo));
            message.Subject = "Test Email Subject";                    

            // Agregar Archivo al Mensaje
            var builder = new BodyBuilder();
            builder.TextBody = emailBody;
            builder.Attachments.Add("Excel.xlsx");

            message.Body = builder.ToMessageBody();
            
            // ->Enviar email
            try
            {
                using var smtp = new SmtpClient();
                smtp.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                smtp.Authenticate(emailFrom, passwordFrom);
                smtp.Send(message);
                smtp.Disconnect(true);
                smtp.Dispose();

                Console.WriteLine("Mensaje enviado!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
            }
        }
    }
}