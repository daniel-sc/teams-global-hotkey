using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using ConsoleHotKey;
using System.Text.RegularExpressions;

[DllImport("user32.dll")]
static extern bool EnumThreadWindows(int dwThreadId, EnumThreadDelegate lpfn, IntPtr lParam);

[DllImport("user32.dll", CharSet = CharSet.Auto)]
static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, int wParam, StringBuilder lParam);

[DllImport("user32.dll")]
static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

[DllImport("user32.dll")]
static extern bool SetForegroundWindow(IntPtr hWnd);

[DllImport("user32.dll")]
static extern IntPtr GetForegroundWindow();

const uint WM_GETTEXT = 0x000D;
const uint WM_KEYDOWN = 0x100;
const uint WM_KEYUP = 0x0101;


/*
Process[] processlist = Process.GetProcesses();

foreach (Process process in processlist)
{
    if (!String.IsNullOrEmpty(process.MainWindowTitle))
    {
        Console.WriteLine("Process: {0} ID: {1} Window title: {2}", process.ProcessName, process.Id, process.MainWindowTitle);
    }
}*/



HotKeyManager.RegisterHotKey(Keys.M, KeyModifiers.Alt | KeyModifiers.Control | KeyModifiers.Shift);
HotKeyManager.HotKeyPressed += new EventHandler<HotKeyEventArgs>(HotKeyManager_HotKeyPressed);
Console.ReadLine();


static void HotKeyManager_HotKeyPressed(object sender, HotKeyEventArgs e)
{
    DateTime start = DateTime.UtcNow;
    Console.WriteLine("finding window..");
    Process? p = findProcess();
    DateTime afterFindProcess = DateTime.UtcNow;

    Console.WriteLine("findProcess: {0} ms", Convert.ToInt32((afterFindProcess - start).TotalMilliseconds));
    Console.WriteLine("\nTeams process: {0}", p?.Id);
    if (p != null)
    {
        IntPtr? w = findCallWindow(p);
        DateTime afterFindWindow = DateTime.UtcNow;

        Console.WriteLine("findWindow: {0} ms", Convert.ToInt32((afterFindWindow - afterFindProcess).TotalMilliseconds));
        if (w != null)
        {
            Console.WriteLine("found call window: {0}", w);

            //Thread.Sleep(200);
            SetForegroundWindow((IntPtr)w);
            //Thread.Sleep(200);
            SendKeys.SendWait("+^M");
            SendKeys.Flush();

            DateTime finish = DateTime.UtcNow;

            Console.WriteLine("sendKeys: {0} ms", Convert.ToInt32((finish - afterFindWindow).TotalMilliseconds));

            /*
            PostMessage((IntPtr)w, WM_KEYDOWN, (IntPtr)(Keys.Control), IntPtr.Zero);
            PostMessage((IntPtr)w, WM_KEYDOWN, (IntPtr)(Keys.E), IntPtr.Zero);
            Thread.Sleep(500);
            PostMessage((IntPtr)w, WM_KEYUP, (IntPtr)(Keys.E), IntPtr.Zero);
            PostMessage((IntPtr)w, WM_KEYUP, (IntPtr)(Keys.Control), IntPtr.Zero);
            */

            // TODO switch back to original window?
        }
        else
        {
            Console.WriteLine("no teams window found!");
        }
    }
    else
    {
        Console.WriteLine("no teams process found!");
    }
}


static Process? findProcess() => Process.GetProcesses()
                  .Where(p => p.ProcessName == "Teams")
                  .FirstOrDefault();

static IEnumerable<IntPtr> EnumerateProcessWindowHandles(int processId)
{
    var handles = new List<IntPtr>();

    foreach (ProcessThread thread in Process.GetProcessById(processId).Threads)
        EnumThreadWindows(thread.Id,
            (hWnd, lParam) => { handles.Add(hWnd); return true; }, IntPtr.Zero);

    return handles;
}

/// returns window actual teams window that is in foreground or else any teams window 
static IntPtr? findCallWindow(Process p)
{
    Regex rx = new Regex(@" Microsoft Teams$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    IEnumerable<IntPtr> windows = EnumerateProcessWindowHandles(p.Id);

    /*foreach (IntPtr windowHandle in windows)
    {
        StringBuilder message = new StringBuilder(1000);
        SendMessage(windowHandle, WM_GETTEXT, message.Capacity, message);
        if (message.Length > 0 && message.ToString().Length > 0)
            Console.WriteLine("{0}: '{1}'", windowHandle, message);
    } */
    IEnumerable<IntPtr> teamsWindows = windows.Where(w =>
    {
        StringBuilder message = new StringBuilder(1000);
        SendMessage(w, WM_GETTEXT, message.Capacity, message);
        //Console.WriteLine("Considering {0}: '{1}'", w, message);
        return rx.IsMatch(message.ToString());

    });
    IntPtr foregroundWindow = GetForegroundWindow();
    //Console.WriteLine("ForgroundWindow: {0}", foregroundWindow);
    IntPtr teamsForegroundWindow = teamsWindows.Where(w => w == foregroundWindow).FirstOrDefault(IntPtr.Zero);
    if (teamsForegroundWindow != IntPtr.Zero)
    {
        return teamsForegroundWindow;
    }
    else
    {
        return teamsWindows.Any() ? teamsWindows.First() : null;
    }

}

delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);
