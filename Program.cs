using System;
using System.Net.Http;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text.RegularExpressions;

class Program : Form
{
    Label statusLabel;
    Label countLabel;
    TextBox logBox;
    ListBox downloadList;
    ProgressBar progressBar;
    Button helpButton;

    HashSet<string> downloaded = new HashSet<string>();
    HashSet<string> loggedNames = new HashSet<string>();

    int downloadCount = 0;

    ClientWebSocket ws;

    string lastFileName = "";

    public Program()
    {
        Text = "FMS e-보고서 수집기 v1.3";
        Width = 380;
        Height = 430;

        statusLabel = new Label();
        statusLabel.Text = "상태 : 시작 준비";
        statusLabel.Top = 10;
        statusLabel.Left = 10;
        statusLabel.Width = 200;

        countLabel = new Label();
        countLabel.Text = "다운로드 : 0";
        countLabel.Top = 10;
        countLabel.Left = 230;
        countLabel.Width = 120;

        progressBar = new ProgressBar();
        progressBar.Top = 35;
        progressBar.Left = 10;
        progressBar.Width = 230;
        progressBar.Minimum = 0;
        progressBar.Maximum = 100;

        helpButton = new Button();
        helpButton.Text = "도움말";
        helpButton.Top = 35;
        helpButton.Left = 250;
        helpButton.Width = 100;
        helpButton.Click += HelpClick;

        logBox = new TextBox();
        logBox.Multiline = true;
        logBox.Top = 70;
        logBox.Left = 10;
        logBox.Width = 340;
        logBox.Height = 120;
        logBox.ScrollBars = ScrollBars.Vertical;

        downloadList = new ListBox();
        downloadList.Top = 200;
        downloadList.Left = 10;
        downloadList.Width = 340;
        downloadList.Height = 170;

        Controls.Add(statusLabel);
        Controls.Add(countLabel);
        Controls.Add(progressBar);
        Controls.Add(helpButton);
        Controls.Add(logBox);
        Controls.Add(downloadList);

        StartProcess();
    }

    void HelpClick(object sender, EventArgs e)
    {
        MessageBox.Show(
    @"FMS e-보고서 수집기 사용 방법

1) .NET 8 Runtime이 설치되어 있어야 합니다.(대부분 Windows 11에는 이미 설치되어 있음)

2) Chrome을 디버그 모드로 실행합니다.(동봉한 Chrome Debug 바로가기로 실행)
2-1) '오류: 대상컴퓨터에서 연결을 거부했으므로 연결하지 못했습니다.'는 Chrome이 Debug모드로 실행되지 않아서 입니다.

3) FMS 로그인 후 e-보고서 열람 화면까지 이동합니다.

4) 'FMS e-보고서 수집기' 프로그램을 실행(이미 실행되어 있으면 껐다가 다시 실행)합니다.

5) 사이트에서 다운받을 파일을 클릭합니다.

6) 프로그램이 PDF 주소를 자동 수집합니다.

7) PDF 파일이 자동 다운로드 됩니다.(프로그램 폴더의 PDF 폴더에 저장)

8) 준공도서가 PDF파일 형식이라면 같은 방법으로 다운이 가능합니다.(TIF 형식은 불가)

제작자: ceryu@kakao.com",
"도움말");
    }

    void Log(string msg)
    {
        if (InvokeRequired)
        {
            Invoke(new Action<string>(Log), msg);
            return;
        }

        logBox.AppendText(msg + Environment.NewLine);
    }

    void AddDownload(string file)
    {
        if (InvokeRequired)
        {
            Invoke(new Action<string>(AddDownload), file);
            return;
        }

        downloadList.Items.Add(file);

        downloadCount++;
        countLabel.Text = "다운로드 : " + downloadCount;
    }

    void UpdateProgress(int value)
    {
        if (InvokeRequired)
        {
            Invoke(new Action<int>(UpdateProgress), value);
            return;
        }

        progressBar.Value = value;
    }

    async void StartProcess()
    {
        try
        {
            statusLabel.Text = "상태 : DevTools 연결중";

            string wsUrl = await GetWebSocketUrl();

            ws = new ClientWebSocket();
            await ws.ConnectAsync(new Uri(wsUrl), CancellationToken.None);

            await Send(ws, @"{""id"":1,""method"":""Network.enable""}");
            await Send(ws, @"{""id"":2,""method"":""Runtime.enable""}");
            await Send(ws, @"{""id"":3,""method"":""Log.enable""}");

            await InjectClickScript();

            statusLabel.Text = "상태 : 감지중";

            byte[] buffer = new byte[1024 * 100];

            while (ws.State == WebSocketState.Open)
            {
                var result = await ws.ReceiveAsync(new ArraySegment<byte>(buffer), CancellationToken.None);

                string msg = Encoding.UTF8.GetString(buffer, 0, result.Count);

                await ProcessMessage(msg);
            }
        }
        catch (Exception ex)
        {
            Log("오류 : " + ex.Message);
        }
    }

    async Task InjectClickScript()
    {
        string script =
        @"{
            ""id"":10,
            ""method"":""Runtime.evaluate"",
            ""params"":{
                ""expression"":""document.addEventListener('click',function(e){var a=e.target.closest('a');if(a){console.log('FMSFILE:'+a.innerText);}});""
            }
        }";

        await Send(ws, script);
    }

    async Task ProcessMessage(string msg)
    {
        if (msg.Contains("FMSFILE:"))
        {
            int i = msg.IndexOf("FMSFILE:");
            string name = msg.Substring(i + 8);

            int end = name.IndexOf("\"");

            if (end > 0)
                name = name.Substring(0, end);

            name = Regex.Unescape(name);

            lastFileName = name.Trim();

            if (!loggedNames.Contains(lastFileName))
            {
                loggedNames.Add(lastFileName);
                Log("파일명 : " + lastFileName);
            }

            UpdateProgress(0);
        }

        if (!msg.Contains("thumbnail"))
            return;

        try
        {
            var json = JObject.Parse(msg);

            var url = json["params"]?["request"]?["url"]?.ToString();

            if (url == null)
                return;

            if (!url.Contains("/renderings/0?resolution=thumbnail"))
                return;

            int idx = url.IndexOf("/renderings");

            string pdfUrl = url.Substring(0, idx);

            if (downloaded.Contains(pdfUrl))
                return;

            downloaded.Add(pdfUrl);

            await DownloadPdf(pdfUrl);
        }
        catch
        {
        }
    }

    async Task DownloadPdf(string url)
    {
        try
        {
            HttpClient client = new HttpClient();

            using (var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead))
            {
                var total = response.Content.Headers.ContentLength ?? -1L;
                var stream = await response.Content.ReadAsStreamAsync();

                string folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PDF");

                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);

                string fileName = lastFileName;

                if (string.IsNullOrWhiteSpace(fileName))
                    fileName = DateTime.Now.Ticks + ".pdf";

                foreach (char c in Path.GetInvalidFileNameChars())
                    fileName = fileName.Replace(c, '_');

                string filePath = Path.Combine(folder, fileName);

                using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    byte[] buffer = new byte[8192];
                    long read = 0;
                    int len;

                    while ((len = await stream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                    {
                        await fs.WriteAsync(buffer, 0, len);
                        read += len;

                        if (total > 0)
                        {
                            int progress = (int)((read * 100) / total);
                            UpdateProgress(progress);
                        }
                    }
                }

                UpdateProgress(100);

                Log("다운로드 완료 : " + fileName);
                AddDownload(fileName);
            }
        }
        catch (Exception ex)
        {
            Log("다운로드 실패 : " + ex.Message);
        }
    }

    async Task Send(ClientWebSocket ws, string msg)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(msg);

        await ws.SendAsync(
            new ArraySegment<byte>(bytes),
            WebSocketMessageType.Text,
            true,
            CancellationToken.None);
    }

    async Task<string> GetWebSocketUrl()
    {
        HttpClient client = new HttpClient();

        string json = await client.GetStringAsync("http://127.0.0.1:9222/json");

        JArray arr = JArray.Parse(json);

        foreach (var item in arr)
        {
            if (item["type"]?.ToString() == "page")
            {
                return item["webSocketDebuggerUrl"].ToString();
            }
        }

        throw new Exception("DevTools 연결 실패");
    }

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.Run(new Program());
    }
}