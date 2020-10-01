using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using log4net;
using Microsoft.Office.Interop.Outlook;
using Quartz;
using Quartz.Impl;

namespace EmailAttachments
{

    class Program
    {
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static void Main(string[] args)
        {
            //EmailReadFromOutLook();
            try
            {
                // construct a scheduler factory
                ISchedulerFactory schedFact = new StdSchedulerFactory();

                // get a scheduler
                IScheduler sched = schedFact.GetScheduler();
                sched.Start();

                IJobDetail job = JobBuilder.Create<EmailReadJob>()
                    .WithIdentity("myJob", "group1")
                    .Build();

                ITrigger firstTrigger = TriggerBuilder
                                    .Create()
                                    .WithIdentity("FirstTrigger").WithDescription("First Job")
                                    // fires 
                                    .WithCronSchedule(ConfigurationManager.AppSettings["FirstInterval"])
                                    // start immediately
                                    .StartAt(DateBuilder.DateOf(DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year))
                                    .Build();

                ITrigger secondTrigger = TriggerBuilder
                                  .Create()
                                  .WithIdentity("SecondTrigger").WithDescription("Second Job")
                                  // fires 
                                  .WithCronSchedule(ConfigurationManager.AppSettings["SecondInterval"])
                                  // start immediately
                                  .StartAt(DateBuilder.DateOf(DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second, DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year))
                                  .Build();


                var dictionary = new Dictionary<IJobDetail, Quartz.Collection.ISet<ITrigger>>();
                dictionary.Add(job, new Quartz.Collection.HashSet<ITrigger>()
                          {
                              firstTrigger,
                              secondTrigger
                          });
                sched.ScheduleJobs(dictionary, true);
            }
            catch (ArgumentException e)
            {
                log.ErrorFormat("Error=", e.Message);
            }

            log.Info("Done");
            //Console.ReadLine();
        }

    }
    public class EmailReadJob : IJob
    {
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public void Execute(IJobExecutionContext context)
        {
            Console.WriteLine("  Started" + context.Trigger.Description);
            log.InfoFormat("{0} - Started", context.Trigger.Description);
            EmailReadFromOutLook();
            log.InfoFormat("{0} - End ", context.Trigger.Description);
            Console.WriteLine("  End " + context.Trigger.Description);
        }

        private static void EmailReadFromOutLook()
        {
            Application outlookApplication = null;
            NameSpace outlookNamespace = null;
            MAPIFolder inboxFolder = null;
            Items mailItems = null;
            log.Info("Emails Reading from OutLook started.");
            try
            {
                outlookApplication = new Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                mailItems = inboxFolder.Items;
                string Filter = "[ReceivedTime] >= Today";

                Items mis = inboxFolder.Items.Restrict(Filter);
                int cnt = mis.Count; ;

                log.InfoFormat("email Count :{0} ", mis.Count);

                string basePath = ConfigurationManager.AppSettings["folderPath"];

                foreach (MailItem item in mis)
                {


                    if (item.Attachments.Count > 0)
                    {
                        foreach (Attachment attach in item.Attachments)
                        {
                            SaveFile(attach);
                        }
                    }
                    //Console.WriteLine(stringBuilder);
                    Marshal.ReleaseComObject(item);
                }
            }
            //Error handler.
            catch (System.Exception e)
            {
                log.ErrorFormat("{0} Exception caught: {0} ", e);
            }
            finally
            {
                ReleaseComObject(mailItems);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }
        }
       
        public static void SaveFile(Attachment attach)
        {
            string basePath = ConfigurationManager.AppSettings["folderPath"];
            var fileName = basePath + attach.FileName;
            string fileExtension = Path.GetExtension(fileName);
            string filenameWithNameWithOutExtension = Path.GetFileNameWithoutExtension(fileName);
            if (fileExtension == "xls" || fileExtension == "xlsx")
            {
                if (!File.Exists(filenameWithNameWithOutExtension + ".csv"))
                {
                    attach.SaveAsFile(filenameWithNameWithOutExtension + ".csv");
                }
                else
                {
                    File.Delete(filenameWithNameWithOutExtension + ".csv");
                }
            }
            else
            {
                if (File.Exists(filenameWithNameWithOutExtension + ".csv"))
                {
                    attach.SaveAsFile(filenameWithNameWithOutExtension + ".csv");
                }
                else
                {
                    File.Delete(filenameWithNameWithOutExtension + ".csv");
                }
            }
        }
        private static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
    }



}
