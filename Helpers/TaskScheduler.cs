using Microsoft.Win32.TaskScheduler;
using System;
using System.Collections.Generic;

namespace EguibarIT.Housekeeping
{
    /// <summary>
    ///
    /// </summary>
    public class TaskScheduler
    {
        /// <summary>
        ///
        /// </summary>
        /// <param name="TaskName"></param>
        /// <param name="ServiceAccount"></param>
        /// <param name="Description"></param>
        /// <param name="PsArguments"></param>
        /// <param name="ExeActionID"></param>
        /// <param name="Ocurrences"></param>
        public static void NewScheduledTask(string TaskName, string ServiceAccount, string Description, string PsArguments, string ExeActionID, EguibarIT.Housekeeping.AdHelper.Ocurrences Ocurrences)
        {
            const string Author = "Eguibar Information Technology S.L.";
            const string Source = "Delegation Model - Housekeeping - " + Author;

            string execute = "C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe";
            string arguments = "-NoLogo -NoProfile -NonInteractive -WindowStyle Minimized ";

            using (TaskService ts = new TaskService())
            {
                TaskFolder tf = ts.RootFolder;

                try
                {
                    // Create a Task Definition
                    TaskDefinition td1 = EguibarIT.Housekeeping.TaskScheduler.NewTask(ts, ServiceAccount, Author, Description, Source);

                    // Create triggers. Value can be One, Two, Three, Four, Six, Eight, Twelve OR Twentyfour each 24 hours
                    TaskScheduler.FullDayTriggers(td1, Ocurrences);

                    // Create an action which opens powershell with parameters
                    td1.Actions.Add(EguibarIT.Housekeeping.TaskScheduler.ExecuteAction(ExeActionID, execute, arguments + PsArguments));

                    // Register the task definition (saves it)
                    //tf.RegisterTaskDefinition("T2 Service Account Housekeeping", td, TaskCreation.CreateOrUpdate, "gMSA_AdTskSchd$", "", TaskLogonType.ServiceAccount, null);
                    tf.RegisterTaskDefinition(TaskName, td1);
                }//end try
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }//end using
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="TaskName"></param>
        /// <param name="ServiceAccount"></param>
        /// <param name="Description"></param>
        /// <param name="PsArguments"></param>
        /// <param name="ExeActionID"></param>
        /// <param name="dicValueQueries"></param>
        /// <param name="EventID"></param>
        /// <param name="Logname"></param>
        public static void NewScheduledEventTask(string TaskName, string ServiceAccount, string Description, string PsArguments, string ExeActionID, Dictionary<string, string> dicValueQueries, string EventID, string Logname)
        {
            const string Author = "Eguibar Information Technology S.L.";
            const string Source = "Delegation Model - Housekeeping - " + Author;

            string execute = "C:\\Windows\\system32\\WindowsPowerShell\\v1.0\\powershell.exe";
            string arguments = "-NoLogo -NoProfile -NonInteractive -WindowStyle Hidden ";

            using (TaskService ts = new TaskService())
            {
                TaskFolder tf;
                TaskFolder _folderExist = ts.GetFolder("Event Viewer Tasks");
                
                if(_folderExist == null)
                {
                    tf = ts.RootFolder.CreateFolder("Event Viewer Tasks");
                }
                else
                {
                    tf = ts.RootFolder.SubFolders["Event Viewer Tasks"];
                }

                

                //tf1 = ts.RootFolder.SubFolders.Any(a => a.Name == "Event Viewer Tasks") ? ts.GetFolder("WarewolfTestFolder") : ts.RootFolder.CreateFolder("WarewolfTestFolder");


                // Create a Task Definition
                TaskDefinition td1 = EguibarIT.Housekeeping.TaskScheduler.NewTask(ts, ServiceAccount, Author, Description, Source);

                // Create triggers based on EventID
                TaskScheduler.EventTriggers(td1, EventID, Logname, dicValueQueries);
                //td1.Triggers.Add(EventTriggers(td1, EventID, Logname, dicValueQueries));

                // Create an action which opens powershell with parameters
                td1.Actions.Add(EguibarIT.Housekeeping.TaskScheduler.ExecuteAction(ExeActionID, execute, arguments + PsArguments));
                
                try
                {
                    // Register the task definition (saves it)
                    //tf.RegisterTaskDefinition("T2 Service Account Housekeeping", td, TaskCreation.CreateOrUpdate, "gMSA_AdTskSchd$", "", TaskLogonType.ServiceAccount, null);
                    tf.RegisterTaskDefinition(TaskName, td1);
                }//end try
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }//end using
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="ts"></param>
        /// <param name="ServiceAccount"></param>
        /// <param name="Author"></param>
        /// <param name="Description"></param>
        /// <param name="Source"></param>
        /// <returns></returns>
        public static TaskDefinition NewTask(TaskService ts, string ServiceAccount, string Author, string Description, string Source)
        {
            // Create a new task definition and assign properties
            TaskDefinition td = ts.NewTask();

            try
            {
                // Account used to execute the task
                td.Principal.UserId = string.Format("{0}\\{1}$", Environment.UserDomainName, ServiceAccount);
                td.Principal.LogonType = TaskLogonType.Password;
                td.Principal.RunLevel = TaskRunLevel.Highest; //.LUA;
                //td.Principal.DisplayName = "DisplayName";

                td.RegistrationInfo.Author = Author;
                td.RegistrationInfo.Description = Description;
                td.RegistrationInfo.Source = Source;
                td.RegistrationInfo.Version = new Version(1, 0);
                //td.RegistrationInfo.URI = new Uri("test://app");
                //td.RegistrationInfo.Documentation = "Don't pretend this is real.";
                td.RegistrationInfo.Date = DateTime.Now;

                td.Settings.AllowDemandStart = true;
                td.Settings.AllowHardTerminate = true;
                td.Settings.Compatibility = TaskCompatibility.V2_3;
                td.Settings.DeleteExpiredTaskAfter = TimeSpan.FromMinutes(5);
                td.Settings.DisallowStartIfOnBatteries = true;
                td.Settings.Enabled = true;
                td.Settings.ExecutionTimeLimit = TimeSpan.FromHours(1);
                td.Settings.Hidden = true;
                td.Settings.IdleSettings.IdleDuration = TimeSpan.FromMinutes(20);
                td.Settings.IdleSettings.RestartOnIdle = false;
                td.Settings.IdleSettings.StopOnIdleEnd = false;
                td.Settings.IdleSettings.WaitTimeout = TimeSpan.FromMinutes(10);
                td.Settings.MultipleInstances = TaskInstancesPolicy.StopExisting;
                td.Settings.Priority = System.Diagnostics.ProcessPriorityClass.Normal;
                td.Settings.RestartCount = 5;
                td.Settings.RestartInterval = TimeSpan.FromSeconds(100);
                td.Settings.RunOnlyIfIdle = false;
                td.Settings.RunOnlyIfNetworkAvailable = false;
                td.Settings.StartWhenAvailable = true;
                td.Settings.StopIfGoingOnBatteries = true;
                td.Settings.WakeToRun = false;

            }//end try
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            return td;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="_random"></param>
        /// <param name="taskID"></param>
        /// <returns></returns>
        public static DailyTrigger DailyTrigger(int _random, string taskID)
        {
            DailyTrigger dt = new DailyTrigger
            {
                StartBoundary = DateTime.Today + TimeSpan.FromMinutes(_random),
                DaysInterval = 1,
                Enabled = true,
                EndBoundary = DateTime.Today + TimeSpan.FromDays(10000),
                ExecutionTimeLimit = TimeSpan.FromMinutes(10),
                RandomDelay = TimeSpan.FromMinutes(30),
                Id = taskID
            };

            return dt;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="id"></param>
        /// <param name="execute"></param>
        /// <param name="arguments"></param>
        /// <returns></returns>
        public static ExecAction ExecuteAction(string id, string execute, string arguments)
        {
            ExecAction Action = new ExecAction
            {
                Id = id,
                Path = execute,
                Arguments = arguments
            };

            return Action;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="td"></param>
        /// <param name="DailyOcurrences"></param>
        public static void FullDayTriggers(TaskDefinition td, EguibarIT.Housekeeping.AdHelper.Ocurrences DailyOcurrences = EguibarIT.Housekeeping.AdHelper.Ocurrences.Twelve)
        {
            const int HoursInDay = 24;
            int NumberOfTriggers;
            int TimeBetweenTriggers;
            int i = 0;

            // Create a random generator
            Random rnd = new Random();

            switch (DailyOcurrences)
            {
                case EguibarIT.Housekeeping.AdHelper.Ocurrences.One:
                    // How many triggers will be created per day
                    NumberOfTriggers = 1;
                    // Time Span between the triggers in Hours
                    TimeBetweenTriggers = HoursInDay / NumberOfTriggers;
                    break;

                case EguibarIT.Housekeeping.AdHelper.Ocurrences.Two:
                    NumberOfTriggers = 2;
                    TimeBetweenTriggers = HoursInDay / NumberOfTriggers;
                    break;

                case EguibarIT.Housekeeping.AdHelper.Ocurrences.Three:
                    NumberOfTriggers = 3;
                    TimeBetweenTriggers = HoursInDay / NumberOfTriggers;
                    break;

                case EguibarIT.Housekeeping.AdHelper.Ocurrences.Four:
                    NumberOfTriggers = 4;
                    TimeBetweenTriggers = HoursInDay / NumberOfTriggers;
                    break;

                case EguibarIT.Housekeeping.AdHelper.Ocurrences.Six:
                    NumberOfTriggers = 6;
                    TimeBetweenTriggers = HoursInDay / NumberOfTriggers;
                    break;

                case EguibarIT.Housekeeping.AdHelper.Ocurrences.Eight:
                    NumberOfTriggers = 8;
                    TimeBetweenTriggers = HoursInDay / NumberOfTriggers;
                    break;

                case EguibarIT.Housekeeping.AdHelper.Ocurrences.Twelve:
                    NumberOfTriggers = 12;
                    TimeBetweenTriggers = HoursInDay / NumberOfTriggers;
                    break;

                case EguibarIT.Housekeeping.AdHelper.Ocurrences.Twentyfour:
                    NumberOfTriggers = 24;
                    TimeBetweenTriggers = HoursInDay / NumberOfTriggers;
                    break;

                case EguibarIT.Housekeeping.AdHelper.Ocurrences.Fortyeight:
                    NumberOfTriggers = 48;
                    TimeBetweenTriggers = HoursInDay / NumberOfTriggers;
                    break;

                default:
                    throw new ArgumentException("You passed in a dodgy value!");
            }

            for (i = 0; i < NumberOfTriggers; i++)
            {
                // Convert TimeBetweenTriggers to minutes
                int minutesBetwennTriggers = TimeBetweenTriggers * 60;
                // Get the random delay in minutes between 0 and the minutesBetwennTriggers minus ramdom delay of tasks start of 30 minutes
                int RandomDelay = rnd.Next(0, (minutesBetwennTriggers - 30));
                // Iterated time span
                int CurrentTrigger = i * minutesBetwennTriggers;

                // Create the current trigger
                td.Triggers.Add(DailyTrigger(CurrentTrigger + RandomDelay, string.Format("Trigger No. {0}", i.ToString())));
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="td"></param>
        /// <param name="EventID"></param>
        /// <param name="Logname"></param>
        /// <param name="dicValueQueries"></param>
        /// <returns></returns>
        public static EventTrigger EventTriggers(TaskDefinition td, string EventID, string Logname, Dictionary<string, string> dicValueQueries)
        {
            EventTrigger eventTrigger = new EventTrigger();

            eventTrigger.Enabled = true;
            eventTrigger.Id = EventID;
            
            eventTrigger.Subscription = @"<QueryList><Query Id='1'><Select Path='Security'>*[System[(EventID='4624')]]</Select></Query></QueryList>";
            //Subscription = "@" + string.Format("\"<QueryList><Query Id='0' Path=\"{0}\"><Select Path=\"{0}\">*[System[(EventID=\"{1}\")]]</Select></Query></QueryList>\"", Logname, EventID),
            //Subscription = string.Format("<QueryList><Query Id='0' Path=\"{0}\"><Select Path=\"{0}\">*[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and EventID=\"{1}\"]]</Select></Query></QueryList>", Logname, EventID),
            //Subscription = string.Format("<QueryList><Query Id='0' Path=\"{0}\"><Select Path=\"{0}\">*[System[(EventID=\"{1}\")]]</Select></Query></QueryList>", Logname, EventID),
            //                      <QueryList><Query Id="0" Path="Security"><Select Path="Security">*[System[(EventID="4624")]]</Select></Query></QueryList>

            eventTrigger.StartBoundary = DateTime.Now;
            eventTrigger.EndBoundary = DateTime.MaxValue;
            eventTrigger.ExecutionTimeLimit = TimeSpan.FromMinutes(10);


            foreach (KeyValuePair<string, string> kvp in dicValueQueries) 
            {
                eventTrigger.ValueQueries.Add(kvp.Key, kvp.Value);
            }

            // Add the newly created EventTrigger to the Task Definition
            td.Triggers.Add(eventTrigger);

            return eventTrigger;
        }
    }//end class
}//end namespace