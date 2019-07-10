using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceProcess;
using Lib.Strings;

namespace Lib.System
{
    public class ServicesFunctions
    {
        public ServiceControllerStatus GetServiceStatus(string serviceName)
        {
            ServiceController sc = new ServiceController(serviceName);

            return sc.Status;
        }

        public string ServiceStatusToString(ServiceControllerStatus serviceControllerStatus)
        { 
            switch (serviceControllerStatus)
            {
                case ServiceControllerStatus.Stopped:
                {
                    return StringsFunctions.ResourceString("resStopped");
                }
                case ServiceControllerStatus.StartPending:
                {
                    return StringsFunctions.ResourceString("resStartPending");
                }
                case ServiceControllerStatus.StopPending:
                {
                    return StringsFunctions.ResourceString("resStopPending");
                }
                case ServiceControllerStatus.Running:
                {
                    return StringsFunctions.ResourceString("resRunning");
                }

                case ServiceControllerStatus.ContinuePending:
                {
                    return StringsFunctions.ResourceString("resContinuePending");
                }
                case ServiceControllerStatus.PausePending:
                {
                    return StringsFunctions.ResourceString("resPausePending");
                }
                case ServiceControllerStatus.Paused:
                {
                    return StringsFunctions.ResourceString("resPaused");
                }

                default:
                {
                    return "Status Changing";
                }
            }
        }

        public string GetServiceStatusAsString(string serviceName)
        {
            return ServiceStatusToString(GetServiceStatus(serviceName));
        }

        public void StopService(string serviceName)
        {
            new Task(() =>
            {
                using (ServiceController sc = new ServiceController(serviceName))
                {
                    sc.Stop();
                }
            }).Start();
        }

        public void StartService(string serviceName)
        {
            new Task(() =>
            {
                using (ServiceController sc = new ServiceController(serviceName))
                {
                    sc.Start();
                }
            }).Start();
        }

        public void RestartService(string serviceName)
        {
            //            StopService(serviceName);
            using (ServiceController sc = new ServiceController(serviceName))
            {
                sc.Stop();
                sc.WaitForStatus(ServiceControllerStatus.Stopped);
            }

            StartService(serviceName);
        }
    }
}
