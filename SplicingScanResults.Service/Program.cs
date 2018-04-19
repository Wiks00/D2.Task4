using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Topshelf;

namespace SplicingScanResults.Service
{
    class Program
    {
        static void Main(string[] args)
        {
            HostFactory.Run(
                serv =>
                {
                    serv.Service(() => new Combiner(ConfigurationManager.AppSettings["ScanerRootPath"]));
                    serv.SetServiceName("SplicingScanResultsService");
                    serv.SetDisplayName("Merging scans into a pdf");
                    serv.StartAutomatically();
                    serv.RunAsLocalService();
                }
            );
        }
    }
}
