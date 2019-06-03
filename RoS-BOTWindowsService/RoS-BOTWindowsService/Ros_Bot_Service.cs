using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.ServiceProcess;
using System.Threading.Tasks;
using System.Timers;

namespace RoS_BOTWindowsService
{
    /// <summary>
    /// RoS-Bot Server By C.Y. Fang
    /// </summary>
    public partial class Ros_Bot_Service : ServiceBase
    {
        /// <summary>
        /// url
        /// </summary>
        private readonly string URL = @"https://www.ros-bot.com";
        /// <summary>
        /// Login url
        /// </summary>
        private readonly string URL_LOGIN = @"https://www.ros-bot.com/user/login";
        /// <summary>
        /// activity url
        /// </summary>
        private readonly string URL_ACTIVITY = @"https://www.ros-bot.com/user/{0}/bot-activity";
        /// <summary>
        /// Bat path
        /// </summary>
        private readonly string BAT_PATH = String.Format(@"{0}\Reset VM.bat", Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
        /// <summary>
        /// Source name
        /// </summary>
        private readonly string SourceName = "RoS-BoT";
        /// <summary>
        /// Log name
        /// </summary>
        private readonly string LogName = "RoS-BoTLog";
        /// <summary>
        /// Config
        /// </summary>
        private readonly Config config = new Config().Deserialize();
        /// <summary>
        /// process info
        /// </summary>
        private readonly ProcessStartInfo startInfo = new ProcessStartInfo
        {
            CreateNoWindow = true,
            UseShellExecute = false,
            WindowStyle = ProcessWindowStyle.Hidden,
            RedirectStandardError = true,
            RedirectStandardOutput = true
        };
        /// <summary>
        /// User id
        /// </summary>
        private string Id { set; get; }
        /// <summary>
        /// Cookie
        /// </summary>
        private CookieContainer Cookie { set; get; }
        /// <summary>
        /// Timer
        /// </summary>
        private System.Timers.Timer timer;


        /// <summary>
        /// Constructor
        /// </summary>
        public Ros_Bot_Service()
        {
            InitializeComponent();

            //確認紀錄是否存在
            if (!EventLog.SourceExists(SourceName))
            {
                //建立Windows紀錄
                EventLog.CreateEventSource(SourceName, LogName);
            }
            //setting source name
            eventLog1.Source = SourceName;
            //setting log name 
            eventLog1.Log = LogName;
            //setting bat path
            startInfo.FileName = BAT_PATH;
            //create bat file to bat path
            CreateBatFile();
        }

        /// <summary>
        /// 建立Bat.exe
        /// </summary>
        private void CreateBatFile()
        {
            try
            {
                if (!File.Exists(BAT_PATH))
                {
                    File.WriteAllLines(BAT_PATH, new string[] {
                    "@echo off",
                    "chcp 65001",
                    String.Format(@"""{0}"" reset ""{1}""", config.VMrunPath, config.MachinePath) });
                }
            }
            catch (FileNotFoundException ex)
            {
                eventLog1.WriteEntry(ex.StackTrace, EventLogEntryType.Warning);
            }
        }

        /// <summary>
        /// 設定使用者ID
        /// </summary>
        /// <param name="content">html content</param>
        private void SetUserID(string content)
        {
            try
            {
                HtmlDocument document = new HtmlDocument();
                document.LoadHtml(content);
                string url = document.DocumentNode.SelectSingleNode(@"//link[@rel='shortlink']").Attributes["href"].Value;
                string id = url.Substring(url.LastIndexOf('/') + 1);
                this.Id = id;
                eventLog1.WriteEntry("ID=" + this.Id);
            }
            catch (ArgumentOutOfRangeException ex)
            {
                eventLog1.WriteEntry(ex.StackTrace, EventLogEntryType.Warning);
            }
            catch (NullReferenceException ex)
            {
                eventLog1.WriteEntry(ex.StackTrace, EventLogEntryType.Warning);
            }
        }

        /// <summary>
        /// 登入RoS-BoT網站
        /// </summary>
        /// <returns>result</returns>
        private async Task<string> LoginPageAsync()
        {
            var cookieContainer = new CookieContainer();
            var handler = new HttpClientHandler
            {
                CookieContainer = cookieContainer,
                UseCookies = true
            };

            string result = string.Empty;
            try
            {
                using (var client = new HttpClient(handler))
                {
                    client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Safari/537.36");
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Origin", URL);
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Referer", URL_LOGIN);
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/x-www-form-urlencoded");
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Accept-Language", "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7");
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Cache-Control", "max-age=0");
                    client.DefaultRequestHeaders.Connection.Add("keep-alive");
                    client.MaxResponseContentBufferSize = int.MaxValue;

                    var formUrlEncodedContent = new FormUrlEncodedContent(new[] {
                    new KeyValuePair<string,string>("form_id","user_login"),
                    new KeyValuePair<string, string>("op", "login"),
                    new KeyValuePair<string, string>("name", config.User),
                    new KeyValuePair<string, string>("pass", config.Password),
                    new KeyValuePair<string, string>("form-build-id", GetFromID())
                });

                    var response = await client.PostAsync(URL_LOGIN, formUrlEncodedContent);
                    var content = response.Content;
                    response.EnsureSuccessStatusCode();
                    this.Cookie = cookieContainer;
                    result = await content.ReadAsStringAsync();
                    SetUserID(result);
                }
            }
            catch (NullReferenceException ex)
            {
                eventLog1.WriteEntry(ex.StackTrace, EventLogEntryType.Warning);
            }
            catch (ArgumentException ex)
            {
                eventLog1.WriteEntry(ex.StackTrace, EventLogEntryType.Warning);
            }

            return result;
        }

        /// <summary>
        /// 取得Bot活動畫面
        /// </summary>
        /// <returns>html content</returns>
        private async Task<string> GetBotActivity()
        {
            var handler = new HttpClientHandler
            {
                CookieContainer = this.Cookie,
                UseCookies = true
            };

            string result = string.Empty;
            try
            {
                using (var client = new HttpClient(handler))
                {
                    client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Safari/537.36");
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Host", URL.Replace("https://", ""));
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "text/html; charset=utf-8");
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Accept-Language", "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7");
                    client.DefaultRequestHeaders.TryAddWithoutValidation("Upgrade-Insecure-Requests", "1");
                    client.DefaultRequestHeaders.Connection.Add("keep-alive");
                    client.MaxResponseContentBufferSize = int.MaxValue;

                    var response = await client.GetAsync(String.Format(URL_ACTIVITY, Id));
                    var content = response.Content;
                    result = await content.ReadAsStringAsync();
                    CheckBotStatus(result);
                }
            }
            catch (HttpRequestException ex)
            {
                eventLog1.WriteEntry(ex.StackTrace, EventLogEntryType.Warning);
            }

            return result;
        }

        /// <summary>
        /// 檢查是否逾時
        /// </summary>
        /// <param name="content">html content</param>
        private void CheckBotStatus(String content)
        {
            HtmlDocument document = new HtmlDocument();
            document.LoadHtml(content);
            string time = document.DocumentNode.SelectSingleNode(@"//small[@class='text-navy']").InnerText;
            int min = int.Parse(config.Timeout) / 60;

            try
            {
                if (time.Contains("hour") || time.Contains("hours"))
                {
                    string[] temp = time.Replace(" ", "").Replace("ago.", "").Replace("hours", ";").Replace("hour", ";").Replace("min", ";").Split(new char[] { ';' });
                    int lastHour = (int.Parse(temp[0]) * 60) + int.Parse(temp[1]);
                    if (lastHour >= min)
                    {
                        eventLog1.WriteEntry(String.Format("1.已經超過了{0}分鐘尚未動作", lastHour - min));
                        ResetVM();
                    }
                }
                else if (time.Contains("min"))
                {
                    string[] temp = time.Replace(" ", "").Replace("ago.", "").Replace("min", ";").Split(new char[] { ';' });
                    int lastMin = int.Parse(temp[0]);
                    if (lastMin >= min)
                    {
                        eventLog1.WriteEntry(String.Format("2.已經超過了{0}分鐘尚未動作", lastMin - min));
                        ResetVM();
                    }
                }
            }
            catch (ArgumentException ex)
            {
                eventLog1.WriteEntry(ex.StackTrace, EventLogEntryType.Warning);
            }
            catch (FormatException ex)
            {
                eventLog1.WriteEntry(ex.StackTrace, EventLogEntryType.Warning);
            }
        }

        /// <summary>
        /// 重啟VM
        /// </summary>
        private void ResetVM()
        {
            try
            {
                using (Process process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                }
            }
            catch (Win32Exception ex)
            {
                eventLog1.WriteEntry(ex.StackTrace, EventLogEntryType.Warning);
            }
        }

        /// <summary>
        /// 取得使用者ID
        /// </summary>
        /// <returns>使用者ID</returns>
        private String GetFromID()
        {
            HtmlWeb web = new HtmlWeb();
            HtmlDocument document = web.Load(URL_LOGIN);
            string value = document.DocumentNode.SelectSingleNode(@"//input[@type='hidden'][@name='form_build_id']").Attributes["value"].Value;
            return value;
        }

        /// <summary>
        /// 服務啟動時
        /// </summary>
        /// <param name="args">args</param>
        protected override void OnStart(string[] args)
        {
            eventLog1.WriteEntry("服務已啟動");
            timer = new System.Timers.Timer();
            //ms=s*1000
            int interval = int.Parse(config.Interval) * 1000;
            timer.Interval = interval;
            timer.Elapsed += Timer_Elapsed;
            timer.Start();
            if (Cookie == null)
            {
                Task.WaitAll(LoginPageAsync());
            }
        }

        /// <summary>
        /// 服務停止時
        /// </summary>
        protected override void OnStop()
        {
            if (timer != null)
            {
                timer.Stop();
                timer.Close();
                timer = null;
                eventLog1.WriteEntry("服務停止");
            }
        }

        /// <summary>
        /// 計時器啟動
        /// </summary>
        /// <param name="sender">sender</param>
        /// <param name="e">event</param>
        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            Task.WaitAll(GetBotActivity());
        }

    }

}
