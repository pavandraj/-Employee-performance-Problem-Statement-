using Microsoft.Extensions.DependencyInjection;
using ReadExcelData.Interface;
using System;

namespace ReadExcelData
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("WelCome to Hackathon");
            ServiceCollection services = new ServiceCollection();
            Console.WriteLine("Starting Reading Excel data.....");
            IServiceProvider serviceProvider = services.BuildServiceProvider();
            var service = serviceProvider.GetService(IReadExcelData);
            service.ReadExcelData();
        }
    }
}
