using System.Management;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System;
using System.Text;
using Microsoft.Extensions.Configuration;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {

            var builder = new ConfigurationBuilder()
                     .AddJsonFile($"appsettings.json", true, true);


            IConfiguration config = builder.Build();


            var smtpHost = config["SMTP:HOST"];
            var smtpPort = config["SMTP:PORT"];
            var smtpUserName = config["SMTP:USERNAME"];
            var smtpPassword = config["SMTP:PASSWORD"];
            var smtpFrom = config["SMTP:FROM"];
            var smtpTo = config["SMTP:TO"];


            StringBuilder sb = new StringBuilder();

            var filename = $"{DateTime.Now.ToString("yyyyMMddHHmmssffff")}.txt";

            Console.WriteLine("--------------- DiagTool - Ferramenta de diagnóstico e analise de hardware. ---------------");

            Console.Write("Digite seu nome completo: ");
            var name = Console.ReadLine();

            Console.Write("Digite seu e-mail: ");
            var email = Console.ReadLine();

            if (name == null)
            {
                name = "n/a";
            }
            if (email == null)
            {
                email = "n/a";
            }


            sb.AppendLine("==========================================================");
            sb.AppendLine(name.ToUpper());
            sb.AppendLine(email.ToUpper());
            sb.AppendLine("==========================================================");

            Console.WriteLine("Iniciando analise do sistema, por favor aguarde...");

            sb.Append("MotherBoard Manufacturer: ");
            GetComponentAndAppend("Win32_BaseBoard", "Manufacturer", sb);

            sb.Append("MotherBoard Model: ");
            GetComponentAndAppend("Win32_BaseBoard", "Product", sb);

            sb.Append("CPU: ");
            GetComponentAndAppend("Win32_Processor", "Name", sb);

            sb.Append("GPU: ");
            GetComponentAndAppend("Win32_VideoController", "Name", sb);

            sb.AppendLine("Ram Memory Installed: ");
            GetComponentAndAppend("Win32_PhysicalMemory", "Capacity", sb);

            sb.Append("BIOS Brand: ");
            GetComponentAndAppend("Win32_BIOS", "Manufacturer", sb);

            sb.Append("BIOS version: ");
            GetComponentAndAppend("Win32_BIOS", "Name", sb);

            sb.Append("Audio:");
            GetComponentAndAppend("Win32_SoundDevice", "ProductName", sb);

            sb.Append("Optical Drives:");
            GetComponentAndAppend("Win32_CDROMDrive", "Name", sb);

            sb.Append("Device Name:");
            GetComponentAndAppend("Win32_ComputerSystem", "Name", sb);

            sb.Append("HDD:");
            GetComponentAndAppend("Win32_DiskDrive", "Model", sb);

            sb.Append("Network:");
            GetComponentAndAppend("Win32_NetworkAdapter", "Name", sb);

            File.AppendAllText(filename, sb.ToString());

            sb.Clear();

            try
            {
                Console.WriteLine("Processando...");
                SmtpClient mySmtpClient = new SmtpClient(smtpHost, Int32.Parse(smtpPort));

                mySmtpClient.UseDefaultCredentials = false;
                System.Net.NetworkCredential basicAuthenticationInfo = new System.Net.NetworkCredential(smtpUserName, smtpPassword);
                mySmtpClient.Credentials = basicAuthenticationInfo;
                mySmtpClient.EnableSsl = true;

                MailAddress from = new MailAddress(smtpFrom, "DiagTool");
                MailAddress to = new MailAddress(smtpTo);
                MailMessage myMail = new System.Net.Mail.MailMessage(from, to);


                // Create  the file attachment for this email message.
                Attachment data = new Attachment(filename, MediaTypeNames.Application.Octet);
                // Add time stamp information for the file.
                ContentDisposition disposition = data.ContentDisposition;
                disposition.CreationDate = System.IO.File.GetCreationTime(filename);
                disposition.ModificationDate = System.IO.File.GetLastWriteTime(filename);
                disposition.ReadDate = System.IO.File.GetLastAccessTime(filename);
                // Add the file attachment to this email message.
                myMail.Attachments.Add(data);

                // add ReplyTo
                MailAddress replyTo = new MailAddress(smtpFrom);
                myMail.ReplyToList.Add(replyTo);

                // set subject and encoding
                myMail.Subject = $"DiagTool - {email}";
                myMail.SubjectEncoding = System.Text.Encoding.UTF8;

                // set body-message and encoding
                myMail.Body = "<p>Ol&aacute;,</p><p>Uma nova analise foi enviada, favor verificar o arquivo em anexo.</p><p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p><p><strong>DiagTool</strong> -<em> powered by</em> Casa de Software CPMIDIAS.</p>"; myMail.BodyEncoding = System.Text.Encoding.UTF8;
                // text or html
                myMail.IsBodyHtml = true;

                Console.WriteLine("Enviando informações...");

                mySmtpClient.Send(myMail);

                Console.WriteLine("Informações enviadas com sucesso, pressione ENTER para terminar.");

                Console.Read();
            }

            catch (SmtpException ex)
            {
                Console.WriteLine(ex.Message);
                //throw new ApplicationException
                //("SmtpException has occured: " + ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                //throw ex;
            }


        }

        private static void GetComponentAndAppend(string hwclass, string syntax, StringBuilder sb)
        {
            ManagementObjectSearcher mos = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM " + hwclass);
            foreach (ManagementObject mj in mos.Get())
            {
                if (Convert.ToString(mj[syntax]) != "")
                {
                    if (hwclass == "Win32_PhysicalMemory")
                    {
                        var value = Convert.ToString(Int64.Parse(Convert.ToString(mj[syntax])) / 1024 / 1024 / 1024);

                        sb.AppendLine($"Ram Slot:|  {value}Gb | ");

                    }
                    else
                    {
                        sb.AppendLine(Convert.ToString(mj[syntax]));
                    }
                }
            }
        }
    }
}