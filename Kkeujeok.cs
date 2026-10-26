// 끄적노트: 메모장처럼 막 적으면 저장되고, AI가 중요·급함으로 알아서 분류한다.
//   Ctrl+E 목록(네 칸, 드래그로 옮기기) · Ctrl+D 지금 줄 끝냄 · Ctrl+, 설정(AI 고르기) · Ctrl+T 항상 위 · Ctrl+L 기록 폴더 · Ctrl+휠 글자 크기
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyTitle("끄적노트")]
[assembly: System.Reflection.AssemblyProduct("끄적노트")]

// 한 줄의 분류. Q: 1 지금 바로, 2 시간 잡고, 3 빨리 해치우기, 4 나중에
class Sort
{
    public int Q;
    public bool Long;
    public bool Manual;   // 사용자가 드래그로 옮긴 것은 AI가 다시 바꾸지 않는다
    public string Day;    // AI가 분류한 날. 날이 바뀌면 급한 정도가 달라지니 다시 분류한다
    public string Due;    // 적힌 날짜(내일까지, 금요일 등)를 AI가 날짜로 바꾼 것. 달력에 쓴다
    public char Repeat;   // 'D' 매일, 'W' 매주, 'M' 이번 달 (네 칸과 체크 목록 둘 다에 나온다), 0 한 번만
    public string Title;  // 목록에 보일 짧은 이름 (AI가 줄임). 메모 원문은 그대로 둔다
    public string Places; // 어디에 보일지: T 네 칸, C 달력 (null이면 예전처럼 둘 다). 되풀이하는 일은 체크 목록에 늘 나온다

    // 되풀이하지 않는 일이 어디에도 안 보이면 잃어버리니 네 칸에 둔다
    public bool InQuad { get { return Places == null || Places.Contains('T') || (Repeat == 0 && !Places.Contains('C')); } }
    public bool InCalendar { get { return Places == null || Places.Contains('C'); } }

    public DateTime? DueDate
    {
        get
        {
            DateTime d;
            return DateTime.TryParseExact(Due ?? "", "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out d) ? d : (DateTime?)null;
        }
    }
}

static class Quad
{
    public static readonly string[] Title = { "", "🔥 지금 바로", "⭐ 시간 잡고", "⚡ 빨리 해치우기", "💤 나중에" };
    public static readonly string[] Hint = { "", "중요하고 급함", "중요한데 안 급함", "급한데 안 중요", "둘 다 아님" };
    public static readonly Color[] Ink = {
        Color.Empty, Color.FromArgb(209, 75, 47), Color.FromArgb(43, 43, 43),
        Color.FromArgb(196, 120, 20), Color.FromArgb(140, 134, 116) };
    public static readonly Color[] Tint = {
        Color.Empty, Color.FromArgb(253, 232, 224), Color.FromArgb(255, 245, 205),
        Color.FromArgb(252, 236, 214), Color.FromArgb(240, 238, 230) };
}

class Kkeujeok : Form
{
    static readonly string Base = AppDomain.CurrentDomain.BaseDirectory;
    static readonly string MemoPath = Path.Combine(Base, "끄적노트.txt");
    static readonly string SortPath = Path.Combine(Base, "분류.txt");
    static readonly string LogDir = Path.Combine(Base, "기록");
    static readonly string BackupDir = Path.Combine(LogDir, "백업");
    static readonly string SettingsPath = Path.Combine(Base, "설정.txt");
    static readonly string CheckPath = Path.Combine(Base, "체크.txt");
    static readonly Encoding Utf8 = new UTF8Encoding(false);

    public static readonly Color Paper = Color.FromArgb(255, 251, 230);
    static readonly Color Ink = Color.FromArgb(43, 43, 43);
    static readonly Color Muted = Color.FromArgb(154, 144, 112);

    readonly RichTextBox box = new RichTextBox();
    readonly Label status = new Label();
    readonly Button listLink = new Button();
    readonly System.Windows.Forms.Timer idle = new System.Windows.Forms.Timer();
    readonly Dictionary<string, string> cfg = new Dictionary<string, string>();
    public readonly Dictionary<string, Sort> Sorts = new Dictionary<string, Sort>();
    readonly Dictionary<string, string> checks = new Dictionary<string, string>();   // 반복하는 일 → 마지막으로 체크한 날
    Font normal, bold, strike;
    bool loading, aiRunning, aiAgain, justFinished;
    int aiCount;
    string undoBlock, undoHead;   // 마지막으로 끝낸 줄 (되돌리기용)
    int undoLine;
    Sort undoSort;
    ListWindow listWindow;
    public event Action Changed = delegate { };

    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, out IntPtr l);
    [DllImport("imm32.dll")] static extern IntPtr ImmGetContext(IntPtr h);
    [DllImport("imm32.dll")] static extern int ImmGetCompositionStringW(IntPtr c, int i, IntPtr buf, int len);
    [DllImport("imm32.dll")] static extern bool ImmReleaseContext(IntPtr h, IntPtr c);
    const int WM_SETREDRAW = 0x0B, EM_GETOLEINTERFACE = 0x043C, GCS_COMPSTR = 0x0008;
    const int tomSuspend = -9999995, tomResume = -9999994;

    Kkeujeok()
    {
        if (!File.Exists(MemoPath) && File.Exists(Path.Combine(Base, "끄적.txt"))) File.Move(Path.Combine(Base, "끄적.txt"), MemoPath);   // 예전 이름에서 옮김
        LoadSettings();
        About = Get("about", "");
        LoadSorts();
        LoadChecks();
        Text = "끄적노트";
        Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        BackColor = Paper;
        StartPosition = FormStartPosition.Manual;
        Bounds = ParseBounds(Get("bounds", ""), 380, 520);
        TopMost = Get("topmost", "0") == "1";

        float size = float.Parse(Get("size", "12"), System.Globalization.CultureInfo.InvariantCulture);
        normal = new Font("맑은 고딕", size);
        bold = new Font("맑은 고딕", size, FontStyle.Bold);
        strike = new Font("맑은 고딕", size, FontStyle.Strikeout);

        box.Dock = DockStyle.Fill;
        box.BorderStyle = BorderStyle.None;
        box.BackColor = Paper;
        box.ForeColor = Ink;
        box.Font = normal;
        box.DetectUrls = false;
        box.ScrollBars = RichTextBoxScrollBars.Vertical;

        var pad = new Panel { Dock = DockStyle.Fill, Padding = new Padding(14, 12, 6, 0), BackColor = Paper };
        pad.Controls.Add(box);

        // 아래: 상태 줄 + 큰 목록 단추 (누르기 쉽게 크게)
        var bar = new Panel { Dock = DockStyle.Bottom, Height = 40, BackColor = Paper, Padding = new Padding(0, 4, 10, 6) };
        status.Dock = DockStyle.Fill;
        status.Padding = new Padding(14, 0, 0, 0);
        status.TextAlign = ContentAlignment.MiddleLeft;
        status.ForeColor = Muted;
        status.BackColor = Paper;
        status.Font = new Font("맑은 고딕", 9);
        listLink.Dock = DockStyle.Right;
        listLink.Width = 110;
        listLink.Text = "목록 열기";
        listLink.Font = new Font("맑은 고딕", 10, FontStyle.Bold);
        listLink.FlatStyle = FlatStyle.Flat;
        listLink.FlatAppearance.BorderSize = 0;
        listLink.BackColor = Color.FromArgb(255, 226, 140);
        listLink.ForeColor = Color.FromArgb(110, 80, 20);
        listLink.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 212, 100);
        listLink.Cursor = Cursors.Hand;
        listLink.TabStop = false;
        listLink.Click += delegate { OpenList(); };
        var gear = new Button
        {
            Dock = DockStyle.Right, Width = 36, Text = "⚙", Font = new Font("맑은 고딕", 10),
            FlatStyle = FlatStyle.Flat, BackColor = Paper, ForeColor = listLink.ForeColor, Cursor = Cursors.Hand, TabStop = false,
        };
        gear.FlatAppearance.BorderSize = 0;
        gear.FlatAppearance.MouseOverBackColor = Color.FromArgb(245, 236, 200);
        gear.Click += delegate { OpenSettings(); };
        var gap = new Panel { Dock = DockStyle.Right, Width = 6, BackColor = Paper };
        bar.Controls.Add(status);
        bar.Controls.Add(gear);
        bar.Controls.Add(gap);
        bar.Controls.Add(listLink);

        Controls.Add(pad);
        Controls.Add(bar);

        loading = true;
        box.Text = File.Exists(MemoPath) ? File.ReadAllText(MemoPath, Utf8).Replace("\r\n", "\n") : "";
        Highlight();
        box.ClearUndo();
        loading = false;
        box.SelectionStart = box.TextLength;
        Show_("Ctrl+E 목록 · Ctrl+D 끝냄");

        idle.Interval = 500;
        idle.Tick += delegate { OnIdle(); };
        // AI 정리는 목록을 열 때만 한다 (타자 칠 때마다 부르면 사용량이 많이 든다)
        box.TextChanged += delegate {
            if (loading) return;
            justFinished = false;
            idle.Stop(); idle.Start();
            Show_("…");
        };
        box.KeyDown += OnKey;
        box.MouseWheel += delegate(object s, MouseEventArgs e) {
            if ((ModifierKeys & Keys.Control) != 0) BeginInvoke((Action)SyncZoom);
        };
        FormClosing += delegate { OnClose(); };
    }

    // ── 저장 ─────────────────────────────────────────────

    void OnIdle()
    {
        if (Composing()) return;   // 한글 조합 중이면 기다린다
        idle.Stop();
        Save();
        Highlight();
        Changed();
        Show_("저장됨 " + DateTime.Now.ToString("HH:mm"));
    }

    bool Composing()
    {
        IntPtr ctx = ImmGetContext(box.Handle);
        if (ctx == IntPtr.Zero) return false;
        int len = ImmGetCompositionStringW(ctx, GCS_COMPSTR, IntPtr.Zero, 0);
        ImmReleaseContext(box.Handle, ctx);
        return len > 0;
    }

    void Save()
    {
        WriteAtomic(MemoPath, box.Text.Replace("\n", "\r\n"));
        // 메모에서 사라진 줄의 분류는 지운다
        var alive = new HashSet<string>(Items());
        foreach (string k in Sorts.Keys.ToList())
            if (!alive.Contains(k)) Sorts.Remove(k);
        SaveSorts();
        bool gone = false;
        foreach (string k in checks.Keys.ToList())
            if (!alive.Contains(k)) { checks.Remove(k); gone = true; }
        if (gone) SaveChecks();
    }

    static void WriteAtomic(string path, string text)
    {
        string tmp = path + ".tmp";
        File.WriteAllText(tmp, text, Utf8);
        if (File.Exists(path)) File.Replace(tmp, path, null);
        else File.Move(tmp, path);
    }

    static bool Indented(string line) { return line.Length > 0 && char.IsWhiteSpace(line[0]) && line.Trim().Length > 0; }

    // 메모의 할 일들. 들여쓴 줄은 바로 윗줄에 딸린 설명이라 따로 세지 않는다.
    // (빈 줄 빼고, 앞뒤 공백 없이, 겹치면 하나만)
    public List<string> Items()
    {
        Dictionary<string, List<string>> details;
        return Parse(out details);
    }

    List<string> Parse(out Dictionary<string, List<string>> details)
    {
        details = new Dictionary<string, List<string>>();
        var list = new List<string>();
        string head = null;
        foreach (string l in box.Lines)
        {
            string t = l.Trim();
            if (t.Length == 0) continue;
            if (Indented(l) && head != null) { details[head].Add(t); continue; }
            head = t;
            if (!details.ContainsKey(t)) { details[t] = new List<string>(); list.Add(t); }
        }
        return list;
    }

    // ── AI 분류 ──────────────────────────────────────────

    // 아직 분류 안 된 할 일(또는 어제 AI가 분류한 것)을 AI에게 묻는다. 목록을 열 때 부른다.
    public void Classify()
    {
        if (aiRunning) { aiAgain = true; return; }
        string today = DateTime.Today.ToString("yyyy-MM-dd");
        Dictionary<string, List<string>> details;
        var context = Parse(out details);
        var todo = context.Where(t => NeedsSort(t, today)).ToList();
        if (todo.Count == 0) { Show_("정리할 새 줄 없음"); return; }
        // 설명 줄이 있으면 같이 보여 준다
        var asked = todo.Select(t => details[t].Count == 0 ? t : t + "  (설명: " + string.Join(" / ", details[t]) + ")").ToList();

        aiRunning = true;
        aiCount = todo.Count;
        Show_("AI가 " + aiCount + "개 정리하는 중…");
        listLink.Text = "정리 중…";
        if (listWindow != null) listWindow.SetBusy(true, aiCount);
        string command = AiCommand;
        ThreadPool.QueueUserWorkItem(delegate {
            Dictionary<int, Sort> result = null;
            string error = null;
            try { result = AskAI(command, asked, context); }
            catch (Exception ex) { error = ex.Message; }
            try { BeginInvoke((Action)delegate { ClassifyDone(todo, result, error, today); }); }
            catch (InvalidOperationException) { }   // 창이 이미 닫힘
        });
    }

    bool NeedsSort(string t, string today)
    {
        Sort s;
        if (!Sorts.TryGetValue(t, out s)) return true;
        return !s.Manual && s.Day != today;
    }

    void ClassifyDone(List<string> todo, Dictionary<int, Sort> result, string error, string today)
    {
        aiRunning = false;
        listLink.Text = "목록 열기";
        if (result != null)
        {
            int got = 0;
            for (int i = 0; i < todo.Count; i++)
            {
                Sort s;
                if (!result.TryGetValue(i + 1, out s)) continue;
                Sort old;
                if (Sorts.TryGetValue(todo[i], out old) && old.Manual) continue;
                // "내일까지"는 처음 적은 날 기준이다. 다음 날 다시 물어도 날짜는 처음 것을 둔다
                if (old != null && old.Due != null) s.Due = old.Due;
                s.Day = today;
                Sorts[todo[i]] = s;
                got++;
            }
            SaveSorts();
            Highlight();
            Show_("AI 정리 끝 (" + got + "개)");
        }
        else Show_("AI 정리 실패 (⚙ 확인): " + Shorten(error, 40));
        if (listWindow != null) listWindow.SetBusy(false, 0);
        Changed();
        if (aiAgain) { aiAgain = false; Classify(); }
    }

    public static string Prompt(List<string> todo, List<string> context)
    {
        var sb = new StringBuilder();
        DateTime d = DateTime.Today;
        string[] week = { "일", "월", "화", "수", "목", "금", "토" };
        sb.Append("너는 할 일 분류기다. 사용자가 메모장에 막 적은 할 일을 분류한다.\n");
        if (About.Trim().Length > 0) sb.Append("사용자 소개 (무엇이 중요한지 판단할 때 참고): ").Append(About.Trim()).Append('\n');
        sb.Append("오늘은 ").Append(d.ToString("yyyy-MM-dd")).Append(" (").Append(week[(int)d.DayOfWeek]).Append(")이다.\n\n");
        sb.Append("칸:\n1 = 중요하고 급함 (오늘·내일 마감, 미루면 큰일)\n2 = 중요한데 안 급함 (공부·진로·일·건강처럼 결과에 영향이 크지만 마감이 멀거나 없음)\n");
        sb.Append("3 = 급한데 덜 중요 (곧 해야 하지만 짧고 영향이 작음: 연락, 신청, 소소한 처리)\n4 = 둘 다 아님 (언젠가 할 것)\n");
        sb.Append("0 = 할 일이 아님 (생각, 아이디어, 감상, 기록처럼 할 행동이 없는 메모)\n");
        sb.Append("한 번에 1시간 넘게 걸릴 일이면 뒤에 L을 붙인다.\n");
        sb.Append("되풀이해서 하는 일이면 붙인다: 매일 하는 일(매일, 날마다, 하루에 한 번, 꾸준히, 틈틈이, 계속 등)은 D, 매주 하는 일(매주, 주마다, 일주일에 한 번 등)은 W, 매달 하는 일(매달, 한 달에 한 번 등)이나 이번 달 안에만 하면 되는 일(이번 달 안에, 이달 중 등)은 M. 이번 달 안에 한 번만 하면 되는 일은 M과 함께 @이번 달 마지막 날을 붙인다. 한 번만 하는 일에는 붙이지 않는다. 시험·마감 날짜가 있어도 그때까지 꾸준히 하라는 말이 있으면 D와 @날짜를 같이 붙인다.\n");
        sb.Append("날짜나 마감이 적혀 있으면(오늘, 내일, 모레, 금요일, 다음 주 화요일, 11/3 등) 맨 뒤에 @YYYY-MM-DD를 붙인다. 날짜가 안 적혀 있으면 붙이지 않는다.\n\n");
        sb.Append("참고로 메모 전체:\n");
        foreach (string c in context) sb.Append("- ").Append(c).Append('\n');
        sb.Append("\n분류할 줄:\n");
        for (int i = 0; i < todo.Count; i++) sb.Append(i + 1).Append(": ").Append(todo[i]).Append('\n');
        sb.Append("할 일 목록에 보일 짧은 이름도 지어 ' | ' 뒤에 붙인다. 핵심 행동만 15자 안팎으로, 날짜·반복 표현(내일까지, 매일 등)은 이미 따로 표시되니 뺀다. 예: '졸업 전까지 영어 회화 자격증 따야함. 매일 공부할 것' → '영어 회화 공부'. 0번 칸이면 이름은 안 붙여도 된다.\n");
        sb.Append("어디에 보일지도 정해 # 뒤에 붙인다 (되풀이하는 일은 체크 목록에 늘 따로 나온다):\n");
        sb.Append("T = 할 일 목록(네 칸): 시간을 내서 직접 해야 하는 일 (과제, 공부, 준비, 연락, 보고서 쓰기 등)\n");
        sb.Append("C = 달력: 날짜가 정해진 것 (마감, 시험, 약속, 회의, 상담 등)\n");
        sb.Append("약속·회의·상담·시험 보는 것처럼 그 시간에 가기만 하면 되는 일정은 C만 붙인다. 날짜 있는 과제는 TC, 날짜 없는 할 일은 T, 되풀이하는 일은 오늘 해야 할 만큼 중요하면 T, 아니면 아무것도 안 붙여도 된다.\n");
        sb.Append("\n각 줄마다 '번호:칸', 필요하면 L, D/W/M, @날짜, #보일 곳을 그 순서로 붙이고 ' | 짧은 이름'을 붙여 한 줄씩만 답해라.\n예:\n1:1@2026-10-27#TC | 운영체제 과제 제출\n2:1@2026-10-28#C | 진로 상담\n3:2LD#T | 영어 회화 공부\n4:3W | 주간 보고서 보내기\n5:0\n다른 말은 쓰지 마라. 파일이나 도구는 쓰지 마라.");
        return sb.ToString();
    }

    public static Dictionary<int, Sort> AskAI(string command, List<string> todo, List<string> context)
    {
        string output = RunAI(command, Prompt(todo, context));
        var map = new Dictionary<int, Sort>();
        foreach (Match m in Regex.Matches(output, @"(?m)^\W*(\d+)\s*[:：]\s*([0-4])[ \t]*([LDWM]*)[ \t]*(?:@[ \t]*(\d{4}-\d{2}-\d{2}))?[ \t]*(#[ \t]*[TC]*)?[ \t]*(?:\|[ \t]*([^\r\n|]+))?", RegexOptions.IgnoreCase))
        {
            string title = m.Groups[6].Success ? m.Groups[6].Value.Trim() : "";
            var s = new Sort { Q = int.Parse(m.Groups[2].Value), Due = m.Groups[4].Success ? m.Groups[4].Value : null,
                Title = title.Length > 0 ? title : null };
            // #이 있으면 AI가 보일 곳을 정한 것 (# 뒤가 비면 네 칸·달력 어디에도 안 보임)
            if (m.Groups[5].Success) s.Places = m.Groups[5].Value.TrimStart('#').Trim().ToUpperInvariant();
            SetFlags(s, m.Groups[3].Value);
            map[int.Parse(m.Groups[1].Value)] = s;
        }
        if (map.Count == 0) throw new Exception("답을 못 알아들음: " + Shorten(output, 60));
        return map;
    }

    // AI 명령을 실행하고 질문은 표준 입력으로 넘긴다. 답(표준 출력)을 돌려준다.
    public static string RunAI(string command, string prompt)
    {
        string exe, args;
        SplitCommand(command, out exe, out args);
        var psi = new ProcessStartInfo(FindExe(exe), args)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            WorkingDirectory = Path.GetTempPath(),
        };
        using (var p = Process.Start(psi))
        {
            var errTask = p.StandardError.ReadToEndAsync();
            var outTask = p.StandardOutput.ReadToEndAsync();
            byte[] input = Utf8.GetBytes(prompt);
            p.StandardInput.BaseStream.Write(input, 0, input.Length);
            p.StandardInput.Close();
            if (!p.WaitForExit(120000)) { try { p.Kill(); } catch { } throw new Exception("시간 초과"); }
            string output = outTask.Result;
            if (output.Trim().Length == 0) throw new Exception("답이 없음: " + Shorten(errTask.Result, 60));
            return output;
        }
    }

    static void SplitCommand(string command, out string exe, out string args)
    {
        command = (command ?? "").Trim();
        if (command.StartsWith("\""))
        {
            int end = command.IndexOf('"', 1);
            if (end < 0) end = command.Length;
            exe = command.Substring(1, end - 1);
            args = end + 1 < command.Length ? command.Substring(end + 1).Trim() : "";
        }
        else
        {
            int sp = command.IndexOf(' ');
            exe = sp < 0 ? command : command.Substring(0, sp);
            args = sp < 0 ? "" : command.Substring(sp + 1).Trim();
        }
        if (exe.Length == 0) throw new Exception("AI 명령이 비어 있음 (⚙ 설정)");
    }

    // PATH에서 프로그램을 찾는다 (npm으로 깐 codex·gemini는 .cmd)
    public static string FindExe(string name)
    {
        if (Path.IsPathRooted(name)) return name;
        var dirs = (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';').ToList();
        string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        dirs.Add(Path.Combine(home, @".local\bin"));
        dirs.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "npm"));
        foreach (string dir in dirs)
        {
            try
            {
                foreach (string ext in new[] { "", ".exe", ".cmd", ".bat" })
                {
                    string p = Path.Combine(dir.Trim(), name + ext);
                    if (ext.Length == 0 && !Path.HasExtension(name)) continue;
                    if (File.Exists(p)) return p;
                }
            }
            catch (ArgumentException) { }
        }
        throw new Exception("'" + name + "'을(를) 찾지 못함. 설치·로그인했는지 확인 (⚙ 설정)");
    }

    public string AiCommand
    {
        get { return Get("ai_cmd", AiPreset.All[0].Command); }
        set { cfg["ai_cmd"] = value; }
    }
    public string View   // 목록 창에서 마지막으로 본 것: quad, calendar, check
    {
        get { return Get("view", "quad"); }
        set { cfg["view"] = value; }
    }
    // ⚙ 설정의 '나에 대해'. AI가 중요도를 판단할 때 참고한다
    public static string About = "";

    public void SetAbout(string text)
    {
        About = text.Replace("\r", " ").Replace("\n", " ").Trim();
        cfg["about"] = About;
    }

    public string AiName
    {
        get { return Get("ai", AiPreset.All[0].Name); }
        set { cfg["ai"] = value; }
    }

    void OpenSettings()
    {
        using (var w = new SettingsWindow(this)) w.ShowDialog(this);
    }

    // 사용자가 목록에서 칸을 옮김
    public void SetManual(string item, int q)
    {
        Sort s;
        if (!Sorts.TryGetValue(item, out s)) s = new Sort();
        s.Q = q;
        s.Manual = true;
        Sorts[item] = s;
        SaveSorts();
        Highlight();
        Changed();
    }

    void LoadSorts()
    {
        if (!File.Exists(SortPath)) return;
        // 칸 \t L·D·W (없으면 -) \t a|m \t 분류한 날 \t 날짜 \t 할 일
        foreach (string l in File.ReadAllLines(SortPath, Utf8))
        {
            // 7칸이면 짧은 이름이 들어 있다 (예전 6칸 형식도 읽는다)
            string[] p = l.Split(new[] { '\t' }, 7);
            int q;
            if (p.Length < 6 || !int.TryParse(p[0], out q) || q < 0 || q > 4) continue;
            var s = new Sort { Q = q, Manual = p[2] == "m", Day = p[3], Due = p[4].Length > 0 ? p[4] : null };
            SetFlags(s, p[1]);
            if (p.Length == 7 && p[5].Length > 0) s.Title = p[5];
            Sorts[p[p.Length - 1]] = s;
        }
    }

    // 예: "LD", "LD/TC" (/ 뒤는 보일 곳)
    static void SetFlags(Sort s, string flags)
    {
        flags = flags.ToUpperInvariant();
        int slash = flags.IndexOf('/');
        if (slash >= 0) { s.Places = flags.Substring(slash + 1); flags = flags.Substring(0, slash); }
        s.Long = flags.Contains('L');
        s.Repeat = flags.Contains('D') ? 'D' : flags.Contains('W') ? 'W' : flags.Contains('M') ? 'M' : '\0';
    }

    void SaveSorts()
    {
        var sb = new StringBuilder();
        foreach (var kv in Sorts)
        {
            string flags = (kv.Value.Long ? "L" : "") + (kv.Value.Repeat != 0 ? kv.Value.Repeat.ToString() : "")
                + (kv.Value.Places != null ? "/" + kv.Value.Places : "");
            sb.Append(kv.Value.Q).Append('\t').Append(flags.Length > 0 ? flags : "-").Append('\t')
              .Append(kv.Value.Manual ? "m" : "a").Append('\t').Append(kv.Value.Day ?? "").Append('\t')
              .Append(kv.Value.Due ?? "").Append('\t').Append((kv.Value.Title ?? "").Replace('\t', ' ')).Append('\t')
              .Append(kv.Key).Append("\r\n");
        }
        WriteAtomic(SortPath, sb.ToString());
    }

    // ── 매일·매주 체크 ───────────────────────────────────

    public static DateTime WeekStart(DateTime d) { return d.Date.AddDays(-(((int)d.DayOfWeek + 6) % 7)); }   // 월요일

    // 목록에 보일 이름: AI가 지은 짧은 이름, 없으면 메모 원문
    public string ShortName(string item) { Sort s; return Sorts.TryGetValue(item, out s) && !string.IsNullOrEmpty(s.Title) ? s.Title : item; }

    public bool IsRepeat(string item) { Sort s; return Sorts.TryGetValue(item, out s) && s.Repeat != 0; }

    // 체크한 날 (안 했으면 null). 매일 하는 일은 오늘, 매주 하는 일은 이번 주, 매달 하는 일은 이번 달에 체크했을 때만 '했음'이다.
    // 날짜가 붙은 이번 달 일(이번 달 안에 한 번만)은 한 번 체크하면 계속 '했음'이다.
    public DateTime? DoneOn(string item)
    {
        Sort s; string day; DateTime d;
        if (!Sorts.TryGetValue(item, out s) || s.Repeat == 0 || !checks.TryGetValue(item, out day)
            || !DateTime.TryParseExact(day, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out d)) return null;
        DateTime today = DateTime.Today;
        bool done = s.Repeat == 'D' ? d == today
            : s.Repeat == 'W' ? d >= WeekStart(today)
            : s.Due != null || (d.Year == today.Year && d.Month == today.Month);
        return done ? d : (DateTime?)null;
    }

    public void Toggle(string item)
    {
        bool done = DoneOn(item).HasValue;
        char r = Sorts[item].Repeat;
        string what = r == 'D' ? "오늘" : r == 'W' ? "이번 주" : "이번 달";
        if (done) checks.Remove(item);
        else checks[item] = DateTime.Today.ToString("yyyy-MM-dd");
        SaveChecks();
        Log((done ? "☐ 체크 풀기 " : "☑ " + what + " 했음 ") + item);
        Highlight();
        Changed();
        Show_((done ? "체크 풀기: " : "☑ " + what + " 했음: ") + Shorten(item, 20));
    }

    // 반복 여부를 사용자가 직접 바꿈 (AI가 다시 안 바꾼다)
    public void SetRepeat(string item, char r)
    {
        Sort s;
        if (!Sorts.TryGetValue(item, out s)) s = new Sort { Q = 2 };
        s.Repeat = r;
        s.Manual = true;
        Sorts[item] = s;
        SaveSorts();
        Highlight();
        Changed();
    }

    // 네 칸(T)·달력(C)에 보일지 사용자가 직접 바꿈 (AI가 다시 안 바꾼다)
    public void SetPlace(string item, char place, bool show)
    {
        Sort s;
        if (!Sorts.TryGetValue(item, out s)) return;
        string now = (s.InQuad ? "T" : "") + (s.InCalendar ? "C" : "");
        now = now.Replace(place.ToString(), "");
        if (show) now += place;
        s.Places = now;
        s.Manual = true;
        SaveSorts();
        Highlight();
        Changed();
    }

    void LoadChecks()
    {
        if (!File.Exists(CheckPath)) return;
        // 체크한 날 \t 할 일
        foreach (string l in File.ReadAllLines(CheckPath, Utf8))
        {
            string[] p = l.Split(new[] { '\t' }, 2);
            if (p.Length == 2) checks[p[1]] = p[0];
        }
    }

    void SaveChecks()
    {
        var sb = new StringBuilder();
        foreach (var kv in checks) sb.Append(kv.Value).Append('\t').Append(kv.Key).Append("\r\n");
        WriteAtomic(CheckPath, sb.ToString());
    }

    static void Log(string text)
    {
        Directory.CreateDirectory(LogDir);
        File.AppendAllText(Path.Combine(LogDir, DateTime.Now.ToString("yyyy-MM-dd") + ".txt"),
            DateTime.Now.ToString("HH:mm") + "  " + text + "\r\n", Utf8);
    }

    // ── 색칠 ─────────────────────────────────────────────

    // 분류된 칸에 따라 줄 색을 바꾼다. 색칠은 되돌리기(Ctrl+Z) 기록에 넣지 않는다.
    void Highlight()
    {
        dynamic doc = TextDocument();
        if (doc != null) doc.Undo(tomSuspend);
        SendMessage(box.Handle, WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);
        int selStart = box.SelectionStart, selLen = box.SelectionLength;
        try
        {
            box.SelectAll();
            box.SelectionColor = Ink;
            box.SelectionFont = normal;
            string[] lines = box.Lines;
            for (int i = 0; i < lines.Length; i++)
            {
                Sort s;
                if (Indented(lines[i]) || !Sorts.TryGetValue(lines[i].Trim(), out s)) continue;
                box.Select(box.GetFirstCharIndexFromLine(i), lines[i].Length);
                // 매일·매주 하는 일을 이번에 했으면 흐리게 줄을 긋는다
                if (s.Repeat != 0 && DoneOn(lines[i].Trim()).HasValue) { box.SelectionColor = Muted; box.SelectionFont = strike; continue; }
                box.SelectionColor = s.Q == 0 ? Muted : Quad.Ink[s.Q];
                if (s.Q == 1 || s.Q == 2) box.SelectionFont = bold;
            }
        }
        finally
        {
            box.Select(selStart, selLen);
            SendMessage(box.Handle, WM_SETREDRAW, (IntPtr)1, IntPtr.Zero);
            box.Invalidate();
            if (doc != null) doc.Undo(tomResume);
        }
    }

    object TextDocument()
    {
        try
        {
            IntPtr unk;
            SendMessage(box.Handle, EM_GETOLEINTERFACE, IntPtr.Zero, out unk);
            if (unk == IntPtr.Zero) return null;
            object ole = Marshal.GetObjectForIUnknown(unk);
            Marshal.Release(unk);
            return ole;
        }
        catch { return null; }
    }

    // ── 키 ───────────────────────────────────────────────

    void OnKey(object sender, KeyEventArgs e)
    {
        if (e.Control && e.KeyCode == Keys.D)
        {
            // 매일·매주 하는 일이면 지우지 않고 '했음' 체크만 한다
            int ln = box.GetLineFromCharIndex(box.SelectionStart);
            string[] lines = box.Lines;
            if (ln < lines.Length && !Indented(lines[ln]) && IsRepeat(lines[ln].Trim())) Toggle(lines[ln].Trim());
            else FinishLine(ln);
            e.SuppressKeyPress = true;
        }
        else if (e.Control && e.KeyCode == Keys.Z && justFinished) { UndoFinish(); e.SuppressKeyPress = true; }
        else if (e.Control && e.KeyCode == Keys.E) { OpenList(); e.SuppressKeyPress = true; }
        else if (e.Control && e.KeyCode == Keys.Oemcomma) { OpenSettings(); e.SuppressKeyPress = true; }
        else if (e.Control && e.KeyCode == Keys.T)
        {
            TopMost = !TopMost;
            Show_(TopMost ? "항상 위 켜짐" : "항상 위 꺼짐");
            e.SuppressKeyPress = true;
        }
        else if (e.Control && e.KeyCode == Keys.L)
        {
            Directory.CreateDirectory(LogDir);
            Process.Start(LogDir);
            e.SuppressKeyPress = true;
        }
        // 붙여넣기는 글자만 (서식·그림 없이)
        else if ((e.Control && e.KeyCode == Keys.V) || (e.Shift && e.KeyCode == Keys.Insert))
        {
            if (Clipboard.ContainsText()) box.SelectedText = Clipboard.GetText().Replace("\r\n", "\n");
            e.SuppressKeyPress = true;
        }
    }

    // 목록에서 끝낸 일: 메모에서 그 줄을 찾아 끝낸다 (매일·매주 하는 일은 체크)
    public void Finish(string item, bool remove = false)
    {
        if (IsRepeat(item) && !remove) { Toggle(item); return; }
        string[] lines = box.Lines;
        for (int i = 0; i < lines.Length; i++)
            if (!Indented(lines[i]) && lines[i].Trim() == item) { FinishLine(i); return; }
    }

    // 그 줄을 끝낸다. 할 일 줄이면 아래 들여쓴 설명 줄도 같이 치운다.
    void FinishLine(int line)
    {
        string[] lines = box.Lines;
        if (line >= lines.Length || lines[line].Trim().Length == 0) return;
        string text = lines[line].Trim();
        int last = line;
        if (!Indented(lines[line]))
            while (last + 1 < lines.Length && Indented(lines[last + 1])) last++;

        Directory.CreateDirectory(LogDir);
        DateTime now = DateTime.Now;
        Sort s;
        string tag = Sorts.TryGetValue(text, out s) && s.Q >= 1 ? Quad.Title[s.Q].Substring(0, 2) + " " : "";
        var log = new StringBuilder(now.ToString("HH:mm") + "  ✔ " + tag + text + "\r\n");
        for (int i = line + 1; i <= last; i++) log.Append("         ").Append(lines[i].Trim()).Append("\r\n");
        File.AppendAllText(Path.Combine(LogDir, now.ToString("yyyy-MM-dd") + ".txt"), log.ToString(), Utf8);

        undoBlock = string.Join("\n", lines, line, last - line + 1);
        undoHead = Indented(lines[line]) ? null : text;
        undoLine = line;
        undoSort = s;

        int start = box.GetFirstCharIndexFromLine(line);
        int len = box.GetFirstCharIndexFromLine(last) + lines[last].Length - start;
        if (last < lines.Length - 1) len += 1;                    // 뒤 줄바꿈까지
        else if (line > 0) { start -= 1; len += 1; }              // 마지막 줄이면 앞 줄바꿈까지
        loading = true;
        box.Select(start, len);
        box.SelectedText = "";
        loading = false;
        justFinished = true;
        Save();
        Changed();
        Show_("끝냄: " + Shorten(text, 20) + " · Ctrl+Z 되돌리기");
    }

    public bool CanUndo { get { return undoBlock != null; } }

    // 마지막으로 끝낸 줄을 원래 자리에 되돌려 놓는다
    public void UndoFinish()
    {
        if (undoBlock == null) return;
        string[] lines = box.Lines;
        int pos;
        string ins;
        if (box.TextLength > 0 && undoLine < lines.Length) { pos = box.GetFirstCharIndexFromLine(undoLine); ins = undoBlock + "\n"; }
        else { pos = box.TextLength; ins = (box.TextLength > 0 && !box.Text.EndsWith("\n") ? "\n" : "") + undoBlock; }
        loading = true;
        box.Select(pos, 0);
        box.SelectedText = ins;
        loading = false;
        if (undoHead != null && undoSort != null) Sorts[undoHead] = undoSort;
        string text = undoBlock.Split('\n')[0].Trim();
        Log("↩ 되돌림 " + text);
        undoBlock = null;
        justFinished = false;
        Save();
        Highlight();
        Changed();
        Show_("되돌림: " + Shorten(text, 20));
    }

    // 목록에서 고르면 끄적노트 창에서 그 줄을 보여 준다
    public void Reveal(string item)
    {
        if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
        Activate();
        string[] lines = box.Lines;
        for (int i = 0; i < lines.Length; i++)
            if (!Indented(lines[i]) && lines[i].Trim() == item)
            {
                box.Select(box.GetFirstCharIndexFromLine(i), lines[i].Length);
                box.ScrollToCaret();
                box.Focus();
                return;
            }
    }

    // ── 창 ───────────────────────────────────────────────

    void OpenList()
    {
        idle.Stop();
        Save();
        if (listWindow == null || listWindow.IsDisposed)
        {
            Rectangle lb = ParseBounds(Get("list", ""), 640, 520);
            if (!cfg.ContainsKey("list"))   // 처음엔 끄적노트 창 왼쪽에 붙여서 연다
                lb = new Rectangle(Math.Max(Screen.FromControl(this).WorkingArea.Left, Left - 650), Top, 640, Math.Max(Height, 480));
            listWindow = new ListWindow(this, lb);
            listWindow.FormClosed += delegate {
                Rectangle b = listWindow.WindowState == FormWindowState.Normal ? listWindow.Bounds : listWindow.RestoreBounds;
                cfg["list"] = b.X + "," + b.Y + "," + b.Width + "," + b.Height;
                listWindow = null;
            };
            listWindow.Show();
        }
        else listWindow.Activate();
        listWindow.SetBusy(aiRunning, aiCount);
        Classify();
    }

    void SyncZoom()
    {
        float size = Math.Max(8f, Math.Min(28f, normal.Size * box.ZoomFactor));
        box.ZoomFactor = 1f;
        normal = new Font("맑은 고딕", size);
        bold = new Font("맑은 고딕", size, FontStyle.Bold);
        strike = new Font("맑은 고딕", size, FontStyle.Strikeout);
        box.Font = normal;
        Highlight();
    }

    void OnClose()
    {
        idle.Stop();
        if (listWindow != null) listWindow.Close();
        Save();
        if (box.Text.Trim().Length > 0)
        {
            Directory.CreateDirectory(BackupDir);
            File.WriteAllText(Path.Combine(BackupDir, DateTime.Now.ToString("yyyy-MM-dd") + ".txt"),
                box.Text.Replace("\n", "\r\n"), Utf8);
        }
        Rectangle b = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
        cfg["bounds"] = b.X + "," + b.Y + "," + b.Width + "," + b.Height;
        cfg["topmost"] = TopMost ? "1" : "0";
        cfg["size"] = normal.Size.ToString(System.Globalization.CultureInfo.InvariantCulture);
        SaveSettings();
    }

    public void SaveSettings()
    {
        var sb = new StringBuilder();
        foreach (var kv in cfg) sb.Append(kv.Key).Append('=').Append(kv.Value).Append("\r\n");
        File.WriteAllText(SettingsPath, sb.ToString(), Utf8);
    }

    void LoadSettings()
    {
        if (!File.Exists(SettingsPath)) return;
        foreach (string l in File.ReadAllLines(SettingsPath, Utf8))
        {
            int i = l.IndexOf('=');
            if (i > 0) cfg[l.Substring(0, i)] = l.Substring(i + 1);
        }
    }

    string Get(string key, string def) { string v; return cfg.TryGetValue(key, out v) ? v : def; }

    static Rectangle ParseBounds(string s, int w0, int h0)
    {
        string[] p = s.Split(',');
        int x, y, w, h;
        if (p.Length == 4 && int.TryParse(p[0], out x) && int.TryParse(p[1], out y)
            && int.TryParse(p[2], out w) && int.TryParse(p[3], out h))
        {
            var r = new Rectangle(x, y, w, h);
            foreach (Screen sc in Screen.AllScreens)
                if (sc.WorkingArea.IntersectsWith(r)) return r;
        }
        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        return new Rectangle(wa.Right - w0 - 40, wa.Top + 80, w0, h0);
    }

    static string Shorten(string s, int n)
    {
        s = (s ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
        return s.Length > n ? s.Substring(0, n) + "…" : s;
    }

    void Show_(string msg) { status.Text = (TopMost ? "📌 " : "") + msg; }

    [STAThread]
    static void Main()
    {
        bool first;
        using (new Mutex(true, "Kkeujeok.SingleInstance", out first))
        {
            if (!first) return;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Kkeujeok());
        }
    }
}

// 목록 창: 네 칸, 달력, 매일·매주 체크. 네 칸에서는 드래그로 칸을 옮기고, 더블클릭하면 끝낸다.
// 어디서든 오른쪽 클릭하면 끝냄·끄적노트에서 보기·매일/매주로 바꾸기를 고를 수 있다.
class ListWindow : Form
{
    readonly Kkeujeok main;
    readonly ListBox[] lists = new ListBox[5];
    readonly Label[] heads = new Label[5];
    readonly TableLayoutPanel grid = new TableLayoutPanel();
    readonly CalendarView calendar;
    readonly RoutineView routine;
    readonly LinkLabel undo = new LinkLabel();
    readonly ToolTip tip = new ToolTip { InitialDelay = 300 };
    readonly Label foot = new Label();
    readonly Label banner = new Label();
    readonly System.Windows.Forms.Timer spin = new System.Windows.Forms.Timer { Interval = 400 };
    bool busy;
    int busyCount, dots;
    readonly LinkLabel quadTab = new LinkLabel(), calTab = new LinkLabel(), checkTab = new LinkLabel();
    int hidden, elsewhere;
    Point dragFrom;
    int dragIndex = -1;
    ListBox dragList;

    public ListWindow(Kkeujeok main, Rectangle bounds)
    {
        this.main = main;
        Text = "끄적노트 — 목록";
        Icon = main.Icon;
        BackColor = Kkeujeok.Paper;
        StartPosition = FormStartPosition.Manual;
        Bounds = bounds;
        MinimumSize = new Size(420, 360);
        KeyPreview = true;
        KeyDown += delegate(object s, KeyEventArgs e) {
            if (e.KeyCode == Keys.Escape) Close();
            if (e.KeyCode == Keys.F5) main.Classify();
        };

        // 위: 보기 고르기
        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 34, Padding = new Padding(14, 8, 0, 0), BackColor = Kkeujeok.Paper };
        foreach (var tab in new[] { quadTab, calTab, checkTab })
        {
            tab.AutoSize = true;
            tab.Font = new Font("맑은 고딕", 10);
            tab.LinkBehavior = LinkBehavior.NeverUnderline;
            tab.Margin = new Padding(0, 0, 14, 0);
            top.Controls.Add(tab);
        }
        quadTab.Text = "▦ 네 칸";
        calTab.Text = "📅 달력";
        checkTab.Text = "☑ 체크";
        quadTab.LinkClicked += delegate { ShowView("quad"); };
        calTab.LinkClicked += delegate { ShowView("calendar"); };
        checkTab.LinkClicked += delegate { ShowView("check"); };

        grid.Dock = DockStyle.Fill;
        grid.ColumnCount = 2; grid.RowCount = 2;
        grid.Padding = new Padding(8, 0, 8, 0);
        grid.BackColor = Kkeujeok.Paper;
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
        grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

        for (int q = 1; q <= 4; q++)
        {
            var cell = new Panel { Dock = DockStyle.Fill, Margin = new Padding(6), BackColor = Quad.Tint[q], Padding = new Padding(10, 6, 6, 6) };
            var head = new Label { Dock = DockStyle.Top, Height = 44, Font = new Font("맑은 고딕", 11, FontStyle.Bold), ForeColor = Quad.Ink[q] };
            heads[q] = head;
            var lb = new ListBox
            {
                Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, BackColor = Quad.Tint[q],
                Font = new Font("맑은 고딕", 10.5f), IntegralHeight = false, AllowDrop = true,
                DrawMode = DrawMode.OwnerDrawFixed, ItemHeight = 26, Tag = q,
            };
            lb.DrawItem += DrawItem;
            // 짧은 이름으로 보이니, 마우스를 올리면 메모 원문을 보여 준다
            lb.MouseMove += delegate(object s, MouseEventArgs e) {
                var l = (ListBox)s;
                int i = l.IndexFromPoint(e.Location);
                string text = i >= 0 ? (string)l.Items[i] : "";
                if (tip.GetToolTip(l) != text) tip.SetToolTip(l, text);
            };
            lb.MouseDown += delegate(object s, MouseEventArgs e) {
                var l = (ListBox)s;
                if (e.Button == MouseButtons.Right)
                {
                    int i = l.IndexFromPoint(e.Location);
                    if (i < 0) return;
                    l.SelectedIndex = i;
                    ItemMenu(main, (string)l.Items[i], l, e.Location);
                    return;
                }
                dragList = l; dragFrom = e.Location; dragIndex = l.IndexFromPoint(e.Location);
            };
            lb.MouseMove += delegate(object s, MouseEventArgs e) {
                if (e.Button != MouseButtons.Left || dragIndex < 0 || dragList != s) return;
                if (Math.Abs(e.X - dragFrom.X) + Math.Abs(e.Y - dragFrom.Y) < 6) return;
                string item = (string)dragList.Items[dragIndex];
                dragIndex = -1;
                dragList.DoDragDrop(item, DragDropEffects.Move);
            };
            lb.DragOver += delegate(object s, DragEventArgs e) {
                e.Effect = e.Data.GetDataPresent(typeof(string)) ? DragDropEffects.Move : DragDropEffects.None;
            };
            lb.DragDrop += delegate(object s, DragEventArgs e) {
                string item = (string)e.Data.GetData(typeof(string));
                if (item != null) main.SetManual(item, (int)((ListBox)s).Tag);
            };
            lb.DoubleClick += delegate(object s, EventArgs e) {
                var l = (ListBox)s;
                if (l.SelectedItem != null) main.Finish((string)l.SelectedItem);
            };
            lists[q] = lb;
            cell.Controls.Add(lb);
            cell.Controls.Add(head);
            grid.Controls.Add(cell, (q - 1) % 2, (q - 1) / 2);
        }

        calendar = new CalendarView(main) { Dock = DockStyle.Fill, Visible = false };
        routine = new RoutineView(main) { Dock = DockStyle.Fill, Visible = false };

        // 아래: 안내 + 다시 정리
        var bottom = new Panel { Dock = DockStyle.Bottom, Height = 28, BackColor = Kkeujeok.Paper };
        foot.Dock = DockStyle.Fill;
        foot.Padding = new Padding(16, 0, 0, 0);
        foot.TextAlign = ContentAlignment.MiddleLeft;
        foot.ForeColor = Color.FromArgb(154, 144, 112);
        foot.Font = new Font("맑은 고딕", 9);
        var again = new LinkLabel
        {
            Dock = DockStyle.Right, AutoSize = true, Padding = new Padding(0, 6, 16, 0), Text = "🔄 정리",
            Font = new Font("맑은 고딕", 9), LinkColor = Color.FromArgb(120, 100, 40), LinkBehavior = LinkBehavior.HoverUnderline,
        };
        again.LinkClicked += delegate { main.Classify(); };
        undo.Dock = DockStyle.Right;
        undo.AutoSize = true;
        undo.Padding = new Padding(0, 6, 12, 0);
        undo.Text = "↩ 되돌리기";
        undo.Font = again.Font;
        undo.LinkColor = Color.FromArgb(209, 75, 47);
        undo.LinkBehavior = LinkBehavior.HoverUnderline;
        undo.Visible = false;
        undo.LinkClicked += delegate { main.UndoFinish(); };
        bottom.Controls.Add(foot);
        bottom.Controls.Add(undo);
        bottom.Controls.Add(again);

        // AI가 정리하는 동안 위에 뜨는 띠
        banner.Dock = DockStyle.Top;
        banner.Height = 30;
        banner.Padding = new Padding(16, 0, 0, 0);
        banner.TextAlign = ContentAlignment.MiddleLeft;
        banner.BackColor = Color.FromArgb(255, 224, 178);
        banner.ForeColor = Color.FromArgb(150, 80, 10);
        banner.Font = new Font("맑은 고딕", 10, FontStyle.Bold);
        banner.Visible = false;
        spin.Tick += delegate { dots = (dots + 1) % 4; ShowBanner(); };
        FormClosed += delegate { spin.Stop(); };

        Controls.Add(grid);
        Controls.Add(calendar);
        Controls.Add(routine);
        Controls.Add(banner);
        Controls.Add(top);
        Controls.Add(bottom);

        main.Changed += Fill;
        FormClosed += delegate { main.Changed -= Fill; };
        Activated += delegate { Fill(); };   // 날이 바뀌었으면 체크를 새로 보여 준다
        ShowView(main.View);
        Fill();
    }

    void ShowView(string view)
    {
        if (view != "calendar" && view != "check") view = "quad";
        grid.Visible = view == "quad";
        calendar.Visible = view == "calendar";
        routine.Visible = view == "check";
        foreach (var t in new[] { quadTab, calTab, checkTab })
        {
            bool on = (t == quadTab && view == "quad") || (t == calTab && view == "calendar") || (t == checkTab && view == "check");
            t.LinkColor = on ? Color.FromArgb(43, 43, 43) : Color.FromArgb(160, 150, 120);
            t.Font = new Font("맑은 고딕", 10, on ? FontStyle.Bold : FontStyle.Regular);
        }
        main.View = view;
        UpdateFoot();
    }

    // 오른쪽 클릭 메뉴: 끝냄(또는 체크), 끄적노트에서 보기, 매일·매주·한 번만 바꾸기
    public static void ItemMenu(Kkeujeok main, string item, Control at, Point p)
    {
        var m = new ContextMenuStrip { Font = new Font("맑은 고딕", 9.5f), ShowImageMargin = false, ShowCheckMargin = true };
        Sort s;
        main.Sorts.TryGetValue(item, out s);
        char rep = s != null ? s.Repeat : '\0';
        string first = rep == 0 ? "✔ 끝냄 (목록에서 치움)" : main.DoneOn(item).HasValue ? "☐ 체크 풀기"
            : rep == 'D' ? "☑ 오늘 했음" : rep == 'W' ? "☑ 이번 주 했음" : "☑ 이번 달 했음";
        m.Items.Add(first, null, delegate { main.Finish(item); });
        if (rep != 0) m.Items.Add("✔ 이제 안 함 (목록에서 치움)", null, delegate { main.Finish(item, true); });
        m.Items.Add("✏ 끄적노트에서 보기", null, delegate { main.Reveal(item); });
        m.Items.Add(new ToolStripSeparator());
        foreach (var opt in new[] { new KeyValuePair<char, string>('D', "매일 하는 일"), new KeyValuePair<char, string>('W', "매주 하는 일"), new KeyValuePair<char, string>('M', "이번 달에 할 일"), new KeyValuePair<char, string>('\0', "한 번만 하는 일") })
        {
            char r = opt.Key;
            var mi = new ToolStripMenuItem(opt.Value) { Checked = rep == r };
            mi.Click += delegate { main.SetRepeat(item, r); };
            m.Items.Add(mi);
        }
        // 어디에 보일지 직접 바꾸기 (AI가 정한 것을 덮어씀)
        if (s != null)
        {
            m.Items.Add(new ToolStripSeparator());
            foreach (var opt in new[] { new KeyValuePair<char, string>('T', "할 일 목록(네 칸)에 보이기"), new KeyValuePair<char, string>('C', "달력에 보이기") })
            {
                char c = opt.Key;
                bool on = c == 'T' ? s.InQuad : s.InCalendar;
                var mi = new ToolStripMenuItem(opt.Value) { Checked = on };
                mi.Click += delegate { main.SetPlace(item, c, !on); };
                m.Items.Add(mi);
            }
        }
        m.Closed += delegate { at.BeginInvoke((Action)m.Dispose); };
        m.Show(at, p);
    }

    public void SetBusy(bool on, int count)
    {
        busy = on;
        busyCount = count;
        banner.Visible = on;
        Text = on ? "끄적노트 — 목록 (AI 정리 중…)" : "끄적노트 — 목록";
        if (on) { dots = 0; ShowBanner(); spin.Start(); } else spin.Stop();
        for (int q = 1; q <= 4; q++) lists[q].Invalidate();
    }

    void ShowBanner()
    {
        banner.Text = "⏳ AI가 " + (busyCount > 0 ? busyCount + "개 " : "") + "정리하는 중" + new string('.', dots + 1)
            + "   ❔ 표시는 아직 분류 전";
    }

    void UpdateFoot()
    {
        string memo = (elsewhere > 0 ? " · 일정 " + elsewhere + "개는 달력·체크에" : "") + (hidden > 0 ? " · 메모 " + hidden + "개는 뺌" : "");
        if (calendar.Visible) foot.Text = "날짜가 적힌 일만 나와요 · 더블클릭하면 끄적노트에서 보기 · 오른쪽 클릭: 끝냄 등";
        else if (routine.Visible) foot.Text = "누르면 체크 · 매일 것은 내일, 매주 것은 다음 주 월요일에 다시 빈칸 · 오른쪽 클릭: 바꾸기";
        else foot.Text = "드래그로 옮기기 · 더블클릭: 끝냄 (☐는 체크) · 오른쪽 클릭: 더 보기" + memo;
        undo.Visible = main.CanUndo;
    }

    void Fill()
    {
        var byQ = new List<string>[5];
        for (int q = 0; q <= 4; q++) byQ[q] = new List<string>();
        int repeats = 0, done = 0;
        elsewhere = 0;
        foreach (string item in main.Items())
        {
            Sort s;
            bool known = main.Sorts.TryGetValue(item, out s);
            if (!known) { byQ[4].Add(item); continue; }
            // 매일·매주 하는 일은 체크 목록에 늘 나오고, 네 칸에는 이번에 아직 안 했을 때만 나온다
            if (s.Repeat != 0 && s.Q != 0)
            {
                repeats++;
                if (main.DoneOn(item).HasValue) { done++; continue; }
            }
            // 약속·시험처럼 달력에만 둘 것은 네 칸에서 뺀다
            if (s.Q != 0 && !s.InQuad) { elsewhere++; continue; }
            byQ[s.Q].Add(item);
        }
        checkTab.Text = "☑ 체크" + (repeats > 0 ? "  " + done + "/" + repeats : "");
        hidden = byQ[0].Count;
        for (int q = 1; q <= 4; q++)
        {
            // 오래 걸리는 일을 먼저 (시간을 미리 잡아 두게)
            var items = byQ[q].OrderByDescending(t => { Sort s; return main.Sorts.TryGetValue(t, out s) && s.Long; }).ToList();
            lists[q].BeginUpdate();
            lists[q].Items.Clear();
            foreach (string t in items) lists[q].Items.Add(t);
            lists[q].EndUpdate();
            heads[q].Text = Quad.Title[q] + "  " + items.Count + "\n" + Quad.Hint[q];
        }
        calendar.Invalidate();
        routine.Invalidate();
        UpdateFoot();
    }

    void DrawItem(object sender, DrawItemEventArgs e)
    {
        if (e.Index < 0) return;
        var lb = (ListBox)sender;
        int q = (int)lb.Tag;
        string item = (string)lb.Items[e.Index];
        bool sel = (e.State & DrawItemState.Selected) != 0;
        using (var bg = new SolidBrush(sel ? ControlPaint.Dark(Quad.Tint[q], 0.05f) : Quad.Tint[q]))
            e.Graphics.FillRectangle(bg, e.Bounds);
        Sort s;
        bool known = main.Sorts.TryGetValue(item, out s);
        // 매일·매주·이번 달 하는 일은 체크 칸을 붙이고, 이번에 했으면 흐리게 보여 준다 (더블클릭하면 체크)
        bool repeat = known && s.Repeat != 0;
        bool doneNow = repeat && main.DoneOn(item).HasValue;
        string mark = !known ? "❔ " : repeat ? (doneNow ? "☑ " : "☐ ") : (s.Long ? "⏳ " : "· ");
        string every = !repeat ? "" : s.Repeat == 'D' ? "  매일" : s.Repeat == 'W' ? "  매주" : "  이번 달";
        string tail = (!known && busy ? "  (정리 중)" : "") + every + (known && s.DueDate.HasValue ? "  " + s.DueDate.Value.ToString("M'/'d") : "") + (known && s.Manual ? "  ✋" : "");
        var r = new Rectangle(e.Bounds.X + 2, e.Bounds.Y, e.Bounds.Width - 4, e.Bounds.Height);
        TextRenderer.DrawText(e.Graphics, mark + main.ShortName(item) + tail, lb.Font, r, doneNow ? Color.FromArgb(160, 150, 120) : Color.FromArgb(43, 43, 43),
            TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
    }
}

// 달력 보기: 한 달을 칸으로 그리고, 날짜가 있는 할 일을 그날 칸에 칸 색으로 적는다.
class CalendarView : Control
{
    readonly Kkeujeok main;
    DateTime month = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
    readonly List<KeyValuePair<Rectangle, string>> hits = new List<KeyValuePair<Rectangle, string>>();
    Rectangle prevBtn, nextBtn, todayBtn;
    static readonly string[] Week = { "일", "월", "화", "수", "목", "금", "토" };

    public CalendarView(Kkeujeok main)
    {
        this.main = main;
        DoubleBuffered = true;
        BackColor = Kkeujeok.Paper;
        ResizeRedraw = true;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
        hits.Clear();
        var title = new Font("맑은 고딕", 12, FontStyle.Bold);
        var small = new Font("맑은 고딕", 9);
        var itemFont = new Font("맑은 고딕", 8.5f);
        var ink = Color.FromArgb(43, 43, 43);
        var muted = Color.FromArgb(160, 150, 120);

        // 머리: ◀ 2026년 10월 ▶  오늘
        int x0 = 14, y0 = 4;
        prevBtn = new Rectangle(x0, y0, 24, 26);
        TextRenderer.DrawText(g, "◀", small, prevBtn, muted, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        var titleRect = new Rectangle(x0 + 26, y0, 120, 26);
        TextRenderer.DrawText(g, month.ToString("yyyy년 M월"), title, titleRect, ink, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        nextBtn = new Rectangle(titleRect.Right + 2, y0, 24, 26);
        TextRenderer.DrawText(g, "▶", small, nextBtn, muted, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        todayBtn = new Rectangle(nextBtn.Right + 10, y0, 40, 26);
        TextRenderer.DrawText(g, "오늘", small, todayBtn, Color.FromArgb(120, 100, 40), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);

        // 날짜별 할 일
        var byDay = new Dictionary<DateTime, List<string>>();
        foreach (string item in main.Items())
        {
            Sort s;
            if (!main.Sorts.TryGetValue(item, out s) || s.Q == 0 || !s.DueDate.HasValue || !s.InCalendar) continue;
            List<string> l;
            if (!byDay.TryGetValue(s.DueDate.Value, out l)) byDay[s.DueDate.Value] = l = new List<string>();
            l.Add(item);
        }

        int gx = 14, gy = 40, gw = Width - 28, gh = Height - gy - 6;
        int cw = gw / 7;
        for (int c = 0; c < 7; c++)
            TextRenderer.DrawText(g, Week[c], small, new Rectangle(gx + c * cw, gy, cw, 20),
                c == 0 ? Color.FromArgb(209, 75, 47) : muted, TextFormatFlags.HorizontalCenter);
        gy += 22; gh -= 22;

        DateTime first = month.AddDays(-(int)month.DayOfWeek);
        int rows = (int)Math.Ceiling(((int)month.DayOfWeek + DateTime.DaysInMonth(month.Year, month.Month)) / 7.0);
        int ch = Math.Max(40, gh / rows);
        using (var line = new Pen(Color.FromArgb(232, 224, 196)))
        {
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < 7; c++)
                {
                    DateTime day = first.AddDays(r * 7 + c);
                    var cell = new Rectangle(gx + c * cw, gy + r * ch, cw, ch);
                    bool inMonth = day.Month == month.Month;
                    if (day == DateTime.Today)
                        using (var b = new SolidBrush(Color.FromArgb(255, 238, 170))) g.FillRectangle(b, cell);
                    g.DrawRectangle(line, cell);
                    TextRenderer.DrawText(g, day.Day.ToString(), small, new Rectangle(cell.X + 4, cell.Y + 2, 30, 16),
                        inMonth ? (c == 0 ? Color.FromArgb(209, 75, 47) : ink) : Color.FromArgb(205, 198, 175), TextFormatFlags.Left);

                    List<string> items;
                    if (!byDay.TryGetValue(day, out items)) continue;
                    items = items.OrderBy(t => main.Sorts[t].Q).ToList();
                    int iy = cell.Y + 20, lineH = 17;
                    int fit = Math.Max(0, (cell.Bottom - iy - 2) / lineH);
                    for (int i = 0; i < items.Count && i < fit; i++)
                    {
                        Sort s = main.Sorts[items[i]];
                        bool more = i == fit - 1 && items.Count > fit;
                        var rr = new Rectangle(cell.X + 3, iy, cell.Width - 6, lineH - 2);
                        if (more)
                        {
                            TextRenderer.DrawText(g, "+" + (items.Count - i) + "개 더", itemFont, rr, muted, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
                            break;
                        }
                        using (var b = new SolidBrush(Quad.Tint[s.Q])) g.FillRectangle(b, rr);
                        using (var b = new SolidBrush(Quad.Ink[s.Q])) g.FillRectangle(b, rr.X, rr.Y, 3, rr.Height);
                        TextRenderer.DrawText(g, main.ShortName(items[i]), itemFont, new Rectangle(rr.X + 5, rr.Y, rr.Width - 5, rr.Height), ink,
                            TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
                        hits.Add(new KeyValuePair<Rectangle, string>(rr, items[i]));
                        iy += lineH;
                    }
                }
        }
    }

    protected override void OnMouseClick(MouseEventArgs e)
    {
        if (prevBtn.Contains(e.Location)) { month = month.AddMonths(-1); Invalidate(); }
        else if (nextBtn.Contains(e.Location)) { month = month.AddMonths(1); Invalidate(); }
        else if (todayBtn.Contains(e.Location)) { month = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1); Invalidate(); }
    }

    // 더블클릭은 끄적노트 창에서 그 줄을 보여 주기만 한다 (끝내기는 오른쪽 클릭 메뉴에서)
    protected override void OnMouseDoubleClick(MouseEventArgs e)
    {
        foreach (var h in hits)
            if (h.Key.Contains(e.Location)) { main.Reveal(h.Value); return; }
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        if (e.Button != MouseButtons.Right) return;
        foreach (var h in hits)
            if (h.Key.Contains(e.Location)) { ListWindow.ItemMenu(main, h.Value, this, e.Location); return; }
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        bool hand = prevBtn.Contains(e.Location) || nextBtn.Contains(e.Location) || todayBtn.Contains(e.Location);
        Cursor = hand ? Cursors.Hand : Cursors.Default;
        // 칸이 좁아 잘린 글자는 마우스를 올리면 다 보여 준다
        string over = null;
        foreach (var h in hits) if (h.Key.Contains(e.Location)) { over = h.Value; break; }
        if (over != tipText) { tipText = over; tip.SetToolTip(this, over ?? ""); }
    }

    readonly ToolTip tip = new ToolTip { InitialDelay = 200 };
    string tipText;
}

// 매일·매주 체크: 오늘 할 것과 이번 주 할 것을 체크 칸으로 보여 준다. 누르면 체크/풀기.
// 매일 것은 다음 날, 매주 것은 다음 주 월요일에 다시 빈칸이 된다.
class RoutineView : Control
{
    readonly Kkeujeok main;
    readonly List<KeyValuePair<Rectangle, string>> hits = new List<KeyValuePair<Rectangle, string>>();
    static readonly string[] Week = { "일", "월", "화", "수", "목", "금", "토" };
    static readonly Color Green = Color.FromArgb(76, 150, 90), GreenTint = Color.FromArgb(229, 242, 226);

    public RoutineView(Kkeujeok main)
    {
        this.main = main;
        DoubleBuffered = true;
        BackColor = Kkeujeok.Paper;
        ResizeRedraw = true;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        hits.Clear();
        var title = new Font("맑은 고딕", 11.5f, FontStyle.Bold);
        var small = new Font("맑은 고딕", 9);
        var itemFont = new Font("맑은 고딕", 10.5f);
        var doneFont = new Font("맑은 고딕", 10.5f, FontStyle.Strikeout);
        var ink = Color.FromArgb(43, 43, 43);
        var muted = Color.FromArgb(160, 150, 120);

        var daily = new List<string>();
        var weekly = new List<string>();
        var monthly = new List<string>();
        foreach (string item in main.Items())
        {
            Sort s;
            if (!main.Sorts.TryGetValue(item, out s) || s.Q == 0) continue;
            if (s.Repeat == 'D') daily.Add(item);
            else if (s.Repeat == 'W') weekly.Add(item);
            else if (s.Repeat == 'M') monthly.Add(item);
        }

        DateTime today = DateTime.Today, ws = Kkeujeok.WeekStart(today);
        int x = 16, w = Width - 32, y = 12;
        if (daily.Count == 0 && weekly.Count == 0 && monthly.Count == 0)
        {
            TextRenderer.DrawText(g, "아직 매일·매주·이번 달 할 일이 없어요.\n\n끄적노트에 \"매일 영어 단어 30개\", \"매주 보고서 내기\", \"이번 달 안에 건강검진\"처럼 적으면\nAI가 정리할 때 여기로 모아 줘요.\n\n목록에서 오른쪽 클릭 → '매일 하는 일'로 직접 옮길 수도 있어요.",
                itemFont, new Rectangle(x, y + 30, w, Height - 60), muted, TextFormatFlags.HorizontalCenter | TextFormatFlags.WordBreak);
            return;
        }

        var sections = new[] {
            new { Head = "오늘 할 것", Sub = today.ToString("M'/'d") + " (" + Week[(int)today.DayOfWeek] + ")", Items = daily },
            new { Head = "이번 주 할 것", Sub = ws.ToString("M'/'d") + " ~ " + ws.AddDays(6).ToString("M'/'d"), Items = weekly },
            new { Head = "이번 달 할 것", Sub = today.Month + "월 · " + (DateTime.DaysInMonth(today.Year, today.Month) - today.Day) + "일 남음", Items = monthly },
        };
        foreach (var sec in sections)
        {
            int done = sec.Items.Count(t => main.DoneOn(t).HasValue);
            var headSize = TextRenderer.MeasureText(g, sec.Head, title);
            TextRenderer.DrawText(g, sec.Head, title, new Point(x, y), ink);
            TextRenderer.DrawText(g, sec.Sub, small, new Point(x + headSize.Width + 4, y + 4), muted);
            bool all = sec.Items.Count > 0 && done == sec.Items.Count;
            string count = all ? "다 했어요 🎉 " + done + "/" + sec.Items.Count : done + "/" + sec.Items.Count;
            TextRenderer.DrawText(g, count, small, new Rectangle(x, y + 4, w, 20), all ? Green : muted, TextFormatFlags.Right);
            y += 32;

            if (sec.Items.Count == 0)
            {
                TextRenderer.DrawText(g, "없음", small, new Point(x + 4, y), muted);
                y += 34;
                continue;
            }
            foreach (string item in sec.Items)
            {
                DateTime? on = main.DoneOn(item);
                var row = new Rectangle(x, y, w, 32);
                using (var b = new SolidBrush(on.HasValue ? GreenTint : Color.White)) g.FillRectangle(b, row);
                var boxR = new Rectangle(row.X + 10, row.Y + 8, 16, 16);
                if (on.HasValue)
                {
                    using (var b = new SolidBrush(Green)) g.FillRectangle(b, boxR);
                    using (var pen = new Pen(Color.White, 2.2f))
                        g.DrawLines(pen, new[] { new Point(boxR.X + 3, boxR.Y + 8), new Point(boxR.X + 7, boxR.Y + 12), new Point(boxR.X + 13, boxR.Y + 4) });
                }
                else using (var pen = new Pen(Color.FromArgb(170, 160, 130), 1.5f)) g.DrawRectangle(pen, boxR);

                // 매주 것은 무슨 요일에 했는지 붙인다
                // 시험·마감이 있으면 그날까지라고 적는다
                string tail = !on.HasValue ? "" : sec.Items == weekly ? Week[(int)on.Value.DayOfWeek] + "요일에 함"
                    : sec.Items == monthly ? on.Value.ToString("M'/'d") + "에 함" : "";
                DateTime? due = main.Sorts[item].DueDate;
                if (due.HasValue && due.Value >= today)
                    tail = (tail.Length > 0 ? tail + " · " : "") + due.Value.ToString("M'/'d") + "까지";
                int tailW = tail.Length > 0 ? TextRenderer.MeasureText(g, tail, small).Width + 10 : 0;
                TextRenderer.DrawText(g, main.ShortName(item), on.HasValue ? doneFont : itemFont, new Rectangle(row.X + 36, row.Y, row.Width - 44 - tailW, row.Height),
                    on.HasValue ? muted : ink, TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
                if (tail.Length > 0)
                    TextRenderer.DrawText(g, tail, small, new Rectangle(row.Right - tailW - 6, row.Y, tailW, row.Height), on.HasValue ? Green : Color.FromArgb(196, 120, 20),
                        TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
                hits.Add(new KeyValuePair<Rectangle, string>(row, item));
                y += 36;
            }
            y += 16;
        }
    }

    string At(Point p)
    {
        foreach (var h in hits) if (h.Key.Contains(p)) return h.Value;
        return null;
    }

    protected override void OnMouseClick(MouseEventArgs e)
    {
        string item = At(e.Location);
        if (item == null) return;
        if (e.Button == MouseButtons.Left) main.Toggle(item);
        else if (e.Button == MouseButtons.Right) ListWindow.ItemMenu(main, item, this, e.Location);
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        string over = At(e.Location);
        Cursor = over != null ? Cursors.Hand : Cursors.Default;
        if (over != tipText) { tipText = over; tip.SetToolTip(this, over ?? ""); }
    }

    readonly ToolTip tip = new ToolTip { InitialDelay = 400 };
    string tipText;
}

// AI 고르기. 명령은 질문을 표준 입력으로 받고 답을 표준 출력으로 내야 한다.
class AiPreset
{
    public string Name, Command, Login, Note;

    public static readonly AiPreset[] All = {
        new AiPreset { Name = "Claude Code", Command = "claude -p --model haiku",
            Login = "claude", Note = "Claude 구독으로 동작. 로그인 창에서 /login" },
        new AiPreset { Name = "Codex (ChatGPT)", Command = "codex exec --skip-git-repo-check -",
            Login = "codex login", Note = "설치: npm i -g @openai/codex  → ChatGPT 계정으로 로그인" },
        new AiPreset { Name = "Gemini", Command = "gemini -m gemini-2.5-flash",
            Login = "gemini", Note = "설치: npm i -g @google/gemini-cli  → 처음 실행 때 구글 계정 로그인" },
        new AiPreset { Name = "직접 입력", Command = "",
            Login = "", Note = "질문을 표준 입력으로 받아 답을 출력하는 명령이면 무엇이든 됨" },
    };

    public static AiPreset Find(string name)
    {
        foreach (var p in All) if (p.Name == name) return p;
        return All[All.Length - 1];
    }
}

class SettingsWindow : Form
{
    readonly Kkeujeok main;
    readonly ComboBox pick = new ComboBox();
    readonly TextBox cmd = new TextBox(), about = new TextBox();
    readonly Label note = new Label();
    readonly Label result = new Label();
    readonly Button login = new Button(), test = new Button(), ok = new Button(), cancel = new Button();

    public SettingsWindow(Kkeujeok main)
    {
        this.main = main;
        Text = "끄적노트 — 설정";
        Icon = main.Icon;
        BackColor = Kkeujeok.Paper;
        Font = new Font("맑은 고딕", 9.5f);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(440, 392);

        var head = new Label { Text = "정리에 쓸 AI", Font = new Font("맑은 고딕", 11, FontStyle.Bold), Location = new Point(18, 16), AutoSize = true };
        pick.DropDownStyle = ComboBoxStyle.DropDownList;
        pick.Location = new Point(20, 48); pick.Width = 400;
        foreach (var p in AiPreset.All) pick.Items.Add(p.Name);

        note.Location = new Point(20, 80); note.Size = new Size(400, 36); note.ForeColor = Color.FromArgb(120, 110, 80);

        var cmdLabel = new Label { Text = "실행 명령", Location = new Point(18, 120), AutoSize = true };
        cmd.Location = new Point(20, 142); cmd.Width = 400; cmd.Font = new Font("Consolas", 10);

        login.Text = "로그인 / 인증"; login.Location = new Point(20, 180); login.Size = new Size(120, 30);
        test.Text = "시험해 보기"; test.Location = new Point(148, 180); test.Size = new Size(120, 30);
        result.Location = new Point(20, 216); result.Size = new Size(400, 36);

        // 나에 대해: 어떤 일이 중요한지는 사람마다 다르니 AI에게 알려 준다
        var aboutLabel = new Label { Text = "나에 대해 (AI가 무엇이 중요한지 판단할 때 참고)", Location = new Point(18, 260), AutoSize = true };
        about.Location = new Point(20, 282); about.Size = new Size(400, 52);
        about.Multiline = true; about.ScrollBars = ScrollBars.Vertical;
        about.Text = Kkeujeok.About;

        ok.Text = "저장"; ok.Location = new Point(244, 348); ok.Size = new Size(84, 30);
        cancel.Text = "취소"; cancel.Location = new Point(336, 348); cancel.Size = new Size(84, 30);
        AcceptButton = ok; CancelButton = cancel;

        Controls.AddRange(new Control[] { head, pick, note, cmdLabel, cmd, login, test, result, aboutLabel, about, ok, cancel });

        var cur = AiPreset.Find(main.AiName);
        pick.SelectedItem = cur.Name;
        cmd.Text = main.AiCommand;
        note.Text = cur.Note;

        pick.SelectedIndexChanged += delegate {
            var p = AiPreset.Find((string)pick.SelectedItem);
            if (p.Command.Length > 0) cmd.Text = p.Command;
            note.Text = p.Note;
            login.Enabled = p.Login.Length > 0;
            result.Text = "";
        };
        login.Enabled = cur.Login.Length > 0;
        login.Click += delegate {
            var p = AiPreset.Find((string)pick.SelectedItem);
            // 로그인은 대화형이라 터미널 창을 따로 띄운다
            try { Process.Start("cmd.exe", "/k " + p.Login); }
            catch (Exception ex) { result.Text = ex.Message; }
        };
        test.Click += delegate { RunTest(); };
        ok.Click += delegate {
            main.AiName = (string)pick.SelectedItem;
            main.AiCommand = cmd.Text.Trim();
            main.SetAbout(about.Text);
            main.SaveSettings();
            DialogResult = DialogResult.OK;
        };
    }

    void RunTest()
    {
        string command = cmd.Text.Trim();
        test.Enabled = false;
        result.ForeColor = Color.FromArgb(120, 110, 80);
        result.Text = "물어보는 중…";
        var sw = Stopwatch.StartNew();
        var sample = new List<string> { "내일까지 과제 제출", "언젠가 읽을 책 목록 정리" };
        ThreadPool.QueueUserWorkItem(delegate {
            string msg; bool good;
            try
            {
                var r = Kkeujeok.AskAI(command, sample, sample);
                Sort a, b;
                r.TryGetValue(1, out a); r.TryGetValue(2, out b);
                msg = "✔ 됨 (" + sw.Elapsed.TotalSeconds.ToString("0") + "초) — \"내일까지 과제 제출\" → "
                    + (a != null ? Quad.Title[a.Q] : "?") + ", \"언젠가…\" → " + (b != null ? Quad.Title[b.Q] : "?");
                good = true;
            }
            catch (Exception ex) { msg = "✘ 안 됨: " + ex.Message; good = false; }
            try
            {
                BeginInvoke((Action)delegate {
                    test.Enabled = true;
                    result.ForeColor = good ? Color.FromArgb(40, 120, 60) : Color.FromArgb(209, 75, 47);
                    result.Text = msg;
                });
            }
            catch (InvalidOperationException) { }
        });
    }
}
