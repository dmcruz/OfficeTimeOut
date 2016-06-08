using System.Globalization;
using System.IO;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeTimeOut
{
    public partial class frmTimeOut : Form
    {
        CheckForWorkstationLocking workLock = new CheckForWorkstationLocking();
        public frmTimeOut()
        {
            InitializeComponent();

        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            workLock.SysEventsCheck(null, new SessionSwitchEventArgs(SessionSwitchReason.SessionLock));

            e.Cancel = false;

            base.OnFormClosing(e);
            if (e.CloseReason == CloseReason.WindowsShutDown) return;

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            workLock.Run(this);

            workLock.SysEventsCheck(null, new SessionSwitchEventArgs(SessionSwitchReason.SessionUnlock));
        }
    }


    public class CheckForWorkstationLocking : IDisposable
    {
        private const int REQUIRED_WORK_TOTAL_MINUTES = 585;  // 9 hrs 45 mins
        DateTime unlock;
        int currentDayMsg = 0;
        Timer timer;
        int currentWeek = 0; 
        private TimeSpan tsTargetWorkingTime;

        private SessionSwitchEventHandler sseh;

        private string ThisWeekFileName
        {
            get
            {
                return string.Format("{0}_WK{1}.txt", DateTime.Now.Year, GetWeekOfYear(DateTime.Now));
            }
        }
        public void SysEventsCheck(object sender, SessionSwitchEventArgs e)
        {

            switch (e.Reason)
            {
                case SessionSwitchReason.SessionLock:
                    LogTimeOut();

                    break;

                case SessionSwitchReason.SessionUnlock:

                    unlock = DateTime.Now;

                    if (!File.Exists(ThisWeekFileName))
                    {
                        string timeInEntry = string.Format("{0}|", DateTime.Now);
                        File.WriteAllText(ThisWeekFileName, timeInEntry);
                    }
                    else
                    {
                        string logContent = File.ReadAllText(ThisWeekFileName);
                        string[] lines = logContent.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                        string lastTimeIn = lines[lines.Length - 1].Split("|".ToCharArray())[0];
                        DateTime d = DateTime.Parse(lastTimeIn);

                        if (d.Day != DateTime.Now.Day)
                        {
                            d = DateTime.Now;
                            string timeInEntry = Environment.NewLine + string.Format("{0}|", d);
                            File.AppendAllText(ThisWeekFileName, timeInEntry);
                        }

                        unlock = d;
                    }

                    form.lblTimeUnlock.Text = unlock.ToLongTimeString();

                    var calculatedTimeOut = unlock.Add(new TimeSpan(9, 45, 0));
                    form.lblCalculatedTimeout.Text = calculatedTimeOut.ToLongTimeString();

                    UpdateTimes();

                    timer.Start();

                    break;
            }
        }

        void LogTimeOut()
        {
            string logContent1 = File.ReadAllText(ThisWeekFileName);
            string[] lines1 = logContent1.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            string lastTimeIn1 = lines1[lines1.Length - 1].Split("|".ToCharArray())[0];
            DateTime d1 = DateTime.Parse(lastTimeIn1);

            if (d1.Day == DateTime.Now.Day)
            {
                lines1[lines1.Length - 1] = lastTimeIn1 + "|" + DateTime.Now;

                string overwriteContent = string.Empty;
                foreach (var line in lines1)
                {
                    if (overwriteContent == string.Empty)
                    {
                        overwriteContent += line;
                    }
                    else
                    {
                        overwriteContent += Environment.NewLine + line;
                    }
                }

                File.WriteAllText(ThisWeekFileName, overwriteContent);
            }
        }

        
        void UpdateForecast()
        {
            string logContent = File.ReadAllText(ThisWeekFileName);
            string[] lines = logContent.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            // average
            double sumDailyMinutes = 0;
            int targetSum = (lines.Length) * REQUIRED_WORK_TOTAL_MINUTES;
            for (int i = 0; i < lines.Length - 1; i++)
            {
                var lineTimes = lines[i].Split("|".ToCharArray());
                DateTime timein = DateTime.Parse(lineTimes[0]);
                DateTime timeout = DateTime.Parse(lineTimes[1]);
                double timediff = (timeout - timein).TotalMinutes;
                sumDailyMinutes += timediff;
            }

            tsTargetWorkingTime = TimeSpan.FromMinutes(targetSum - sumDailyMinutes);
            form.lblTargetWorkingTime.Text = string.Format("{0} hour(s) & {1} minute(s)",
                tsTargetWorkingTime.Hours,
                tsTargetWorkingTime.Minutes);


            var targetedTimeOut = unlock.Add(tsTargetWorkingTime);
            form.lblTargetTimeout.Text = targetedTimeOut.ToLongTimeString();
        }

        void UpdateTodayWorkingTime()
        {
            var countdown = DateTime.Now - unlock;
            string countdownMsg = string.Format("{0} hour(s) & {1} minute(s)", countdown.Hours, countdown.Minutes);
            form.lblWorkingTime.Text = countdownMsg;
        }

        void UpdateTimeoutCountdown()
        {
            var countdown = unlock.AddMinutes(REQUIRED_WORK_TOTAL_MINUTES) - DateTime.Now;
            string countdownMsg = string.Format("{0} hour(s) & {1} minute(s)", countdown.Hours, countdown.Minutes);
            form.lblCountdown.Text = countdownMsg;

        }

        void UpdateTimes()
        {
            UpdateTimeoutCountdown();

            UpdateForecast();

            UpdateTodayWorkingTime();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            UpdateTimes();


            var totalMins = DateTime.Now.Subtract(unlock).TotalMinutes;
            if (totalMins >= tsTargetWorkingTime.TotalMinutes)
            {
                if (currentDayMsg == DateTime.Now.Day)
                {
                    if (form.WindowState == FormWindowState.Minimized)
                    {
                        form.WindowState = FormWindowState.Normal;
                    }

                    form.Activate();
                    currentDayMsg = DateTime.Now.Day;
                    timer.Stop();
                    //MessageBox.Show("You can now leave the office");
                }
            }

        }

        frmTimeOut form;
        public void Run(frmTimeOut form)
        {
            timer = new Timer();
            timer.Tick += Timer_Tick;
            timer.Interval = 5000;

            this.form = form;
            sseh = new SessionSwitchEventHandler(SysEventsCheck);
            SystemEvents.SessionSwitch += sseh;
        }


        #region IDisposable Members

        public void Dispose()
        {
            SystemEvents.SessionSwitch -= sseh;
        }

        #endregion

        public static int GetWeekOfYear(DateTime time)
        {
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
            {
                time = time.AddDays(3);
            }
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }
    }
}
