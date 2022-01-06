using System.Management;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Motherboard Manufacturer: ");
            GetComponent("Win32_BaseBoard", "Manufacturer");
            Console.Write("Motherboard Model:");
            GetComponent("Win32_BaseBoard", "Product");
            Console.Write("CPU:");
            GetComponent("Win32_Processor", "Name");
            Console.Write("GPUs:");
            GetComponent("Win32_VideoController", "Name");
            Console.WriteLine("Ram Memory Installed");
            GetComponent("Win32_PhysicalMemory", "Capacity");
            Console.Write("BIOS Brand: ");
            GetComponent("Win32_BIOS", "Manufacturer");
            Console.Write("BIOS version: ");
            GetComponent("Win32_BIOS", "Name");
            Console.Write("Audio:");
            GetComponent("Win32_SoundDevice", "ProductName");
            Console.Write("Optical Drives:");
            GetComponent("Win32_CDROMDrive", "Name");
            Console.Write("Device Name:");
            GetComponent("Win32_ComputerSystem", "Name");
            Console.Write("HDD:");
            GetComponent("Win32_DiskDrive", "Model");
            Console.Write("Network:");
            GetComponent("Win32_NetworkAdapter", "Name");
            Console.Read();

            try
            {

                SmtpClient mySmtpClient = new SmtpClient("my.smtp.exampleserver.net");

                // set smtp-client with basicAuthentication
                mySmtpClient.UseDefaultCredentials = false;
                System.Net.NetworkCredential basicAuthenticationInfo = new
                   System.Net.NetworkCredential("username", "password");
                mySmtpClient.Credentials = basicAuthenticationInfo;

                // add from,to mailaddresses
                MailAddress from = new MailAddress("test@example.com", "TestFromName");
                MailAddress to = new MailAddress("test2@example.com", "TestToName");
                MailMessage myMail = new System.Net.Mail.MailMessage(from, to);

                // add ReplyTo
                MailAddress replyTo = new MailAddress("reply@example.com");
                myMail.ReplyToList.Add(replyTo);

                // set subject and encoding
                myMail.Subject = "Test message";
                myMail.SubjectEncoding = System.Text.Encoding.UTF8;

                // set body-message and encoding
                myMail.Body = "<b>Test Mail</b><br>using <b>HTML</b>.";
                myMail.BodyEncoding = System.Text.Encoding.UTF8;
                // text or html
                myMail.IsBodyHtml = true;

                mySmtpClient.Send(myMail);
            }

            catch (SmtpException ex)
            {
                throw new ApplicationException
                  ("SmtpException has occured: " + ex.Message);
            }
            catch (Exception ex)
            {
                throw ex;
            }


        }



        private static void GetComponent(string hwclass, string syntax)
        {
            ManagementObjectSearcher mos = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM " + hwclass);
            foreach (ManagementObject mj in mos.Get())
            {
                if (Convert.ToString(mj[syntax]) != "")
                {
                    if (hwclass == "Win32_PhysicalMemory")
                    {
                        var value = Convert.ToString(Int64.Parse(Convert.ToString(mj[syntax])) / 1024 / 1024 / 1024);
                        Console.WriteLine($"Ram Memory Slot: {value} Gb");
                    }
                    else
                    {

                        Console.WriteLine(Convert.ToString(mj[syntax]));
                    }
                }
            }
        }
    }
}