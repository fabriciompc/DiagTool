using System.Management;

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
            Console.Write("Ram Memory:");
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


        }



        private static void GetComponent(string hwclass, string syntax)
        {
            ManagementObjectSearcher mos = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM " + hwclass);
            foreach (ManagementObject mj in mos.Get())
            {
                if (Convert.ToString(mj[syntax]) != "")
                    Console.WriteLine(Convert.ToString(mj[syntax]));
            }
        }
    }
}