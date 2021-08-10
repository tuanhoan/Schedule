using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using LichThucTap;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Schedule
{
    public partial class Form1 : Form
    {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Demo Read Google SpeadSheet";
        public Form1()
        {
            InitializeComponent();
            UserCredential credential;
            
            this.Size = new Size(980, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Lịch thực tập Maico Group";


            using (var stream =
                new FileStream("../../../credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "../../../token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "tuananh.maicogroup@gmail.com",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            String spreadsheetId = "1PRXejJDJkk45uPiew0puosYp2I3qqtna_jirz-mq1Qo";
            String range = "A2:J";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);


            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            OneWeek Week = new OneWeek();
            var firstDayOfWeek = DateTime.Now.Date.AddDays(-1 * (int)(DateTime.Now.Date.DayOfWeek - 1) );
             firstDayOfWeek =firstDayOfWeek.AddDays(-7);//Muốn xem từ ngày nào thì chỉnh lại đây
            var lastDayOfWeek = firstDayOfWeek.AddDays(7);
            if (values != null)
            {
                foreach (var item in values)
                {
                    DateTime realDateTime = new DateTime();
                    try
                    {
                        realDateTime = DateTime.ParseExact(item[0].ToString(), "dd/MM/yyyy HH:mm:ss",
                                       System.Globalization.CultureInfo.InvariantCulture);
                    }
                    catch { }
                    try
                    {
                        realDateTime = DateTime.ParseExact(item[0].ToString(), "dd/MM/yyyy H:mm:ss",
                                       System.Globalization.CultureInfo.InvariantCulture);
                    }
                    catch { }
                    if (realDateTime.Date >= firstDayOfWeek.Date)
                    {
                        Test(Week.T2, item, 1);
                        Test(Week.T3, item, 2);
                        Test(Week.T4, item, 3);
                        Test(Week.T5, item, 4);
                        Test(Week.T6, item, 5);
                        Test(Week.T7, item, 6);
                        Test(Week.CN, item, 7);
                    }
                }
            }
            OneDay temp = new OneDay()
            {
                Sang = "Sáng    (8h-12h)",
                Chieu = "Chiều (13h30-17h30)",
                Toi = "Tối      (18h-21h) "
            };
            InitButton("Thứ 2", 120, 50);
            InitButton("Thứ 3", 240, 50);
            InitButton("Thứ 4", 360, 50);
            InitButton("Thứ 5", 480, 50);
            InitButton("Thứ 6", 600, 50);
            InitButton("Thứ 7", 720, 50);
            InitButton("Chủ Nhật", 840, 50);
            Set(temp, 0, 100);
            Set(Week.T2, 120, 100);
            Set(Week.T3, 240, 100);
            Set(Week.T4, 360, 100);
            Set(Week.T5, 480, 100);
            Set(Week.T6, 600, 100);
            Set(Week.T7, 720, 100);
            Set(Week.CN, 840, 100); 
            lblTitle.Text = "Lịch làm từ ngày " + firstDayOfWeek.AddDays(7).ToString("dd/MM/yyy") + " đến ngày: " + lastDayOfWeek.AddDays(7).ToString("dd/MM/yyy");
            lblTitle.TextAlign = ContentAlignment.MiddleCenter;
            Console.WriteLine(lblTitle.Text.Length);
            lblTitle.Location = new Point(this.Width/4, 0);
        }
        void Set(OneDay Day, int x, int y)
        {
            Button btn = new Button()
            {
                Text = Day.Sang,
                Location = new Point(x, y),
                Size = new Size(120, 100),
                TextAlign = ContentAlignment.MiddleCenter
            };
            this.Controls.Add(btn);
            Button btn1 = new Button()
            {
                Text = Day.Chieu,
                Location = new Point(x, y + 100),
                Size = new Size(120, 100),
                TextAlign = ContentAlignment.MiddleCenter
            };
            this.Controls.Add(btn1);
            Button btn2 = new Button()
            {
                Text = Day.Toi,
                Location = new Point(x, y + 200),
                Size = new Size(120, 100),
                TextAlign = ContentAlignment.MiddleCenter
            };
            this.Controls.Add(btn2);
        }
        void InitButton(string Name, int vtx, int vty)
        {
            Button btn = new Button()
            {
                Text = Name,
                Location = new Point(vtx, vty),
                Size = new Size(120, 50)
            };
            this.Controls.Add(btn);
        }
        void Test(OneDay Thu, IList<object> item, int vt)
        {
            var name = item[8].ToString().Split(' ')[item[8].ToString().Split(' ').Count() - 2] + " " + item[8].ToString().Split(' ')[item[8].ToString().Split(' ').Count() - 1] + ", "; 
            if (item[vt].ToString().Contains("Sáng"))
            {
                Thu.Sang += name;
            }
            if (item[vt].ToString().Contains("Chiều"))
            {
                Thu.Chieu += name;
            }
            if (item[vt].ToString().Contains("Tối"))
            {
                Thu.Toi += name;
            }
        }
    }
}
