using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Threading;
using System.Timers;
using System.IO;
using Timer = System.Timers.Timer;
using System.Diagnostics;

class Program
{
    public static Dictionary<string, string> dictionary;
    public static Process[] processList;

    static void Main(string[] args)
    {
        dictionary = getConfig();
        processList = Process.GetProcesses();

        // Create a timer with a ten second interval.
        Timer aTimer = new Timer(10000);
        // Hook up the Elapsed event for the timer. 
        aTimer.Elapsed += StartTimer;
        aTimer.AutoReset = true;
        aTimer.Enabled = true;
    }

    public static void StartTimer(object o, ElapsedEventArgs e)
    {
        GetRDCWindows(dictionary, processList);
    }

    //get RDC windows
    public static void GetRDCWindows(Dictionary<string, string> dictionary, Process[] processlist)
    {
        try
        {
            foreach (KeyValuePair<string, string> rdc in dictionary)

            {
                Console.WriteLine("Checking for:{0} ...", rdc.Key);
                int counter = 0;
                foreach (Process process in processlist)
                {
                    if (!String.IsNullOrEmpty(process.MainWindowTitle))
                    {
                        if (process.MainWindowTitle==rdc.Key)
                        {
                            counter++;
                        }
                        else
                        {

                        }
                    }
                }

                if (counter==1)
                {
                    Console.WriteLine("Window:{0} is already opened!!", rdc.Key);
                }
                else
                {
                    Thread.Sleep(2000);
                    Console.WriteLine("Window:{0} is not opened!!", rdc.Key);
                    Console.WriteLine("Window:{0} Opening...", rdc.Key);
                    System.Diagnostics.Process.Start("mstsc.exe");
                    Thread.Sleep(2000);
                    var RDC = System.Diagnostics.Process.GetProcessesByName("mstsc").FirstOrDefault();
                    AutomationElement RDCElement1 = AutomationElement.FromHandle(RDC.MainWindowHandle)
                                    .FindAll(TreeScope.Subtree, Condition.TrueCondition)
                                    .Cast<AutomationElement>()
                                    .Where(x => x.Current.ClassName=="Edit")
                                    .FirstOrDefault();
                    AutomationElement RDCElement2 = AutomationElement.FromHandle(RDC.MainWindowHandle)
                                    .FindAll(TreeScope.Subtree, Condition.TrueCondition)
                                        .Cast<AutomationElement>()
                                        .Where(x => x.Current.ClassName=="Button"&&
                                        (x.Current.Name=="Connect")||(x.Current.Name=="Yes")).FirstOrDefault();

                    Thread.Sleep(2000);
                    EnterValue(RDCElement1, rdc.Value, rdc.Key);
                    //clickFButton("Connect");
                    Thread.Sleep(2000);
                    // clickFButton("Yes");


                }
                Thread.Sleep(5000);
            }
        }
        catch (Exception ex)
        {

        }
    }

    private static void EnterValue(AutomationElement element, string rdcIP, string rdckey)
    {
        try
        {
            AutomationElement TextboxElement = element.FindAll(TreeScope.Subtree, Condition.TrueCondition)
                                            .Cast<AutomationElement>()
                                            .Where(x => x.Current.ClassName=="Edit")
                                            .FirstOrDefault();
            if (TextboxElement!=null)
            {
                if (TextboxElement.TryGetCurrentPattern(ValuePattern.Pattern, out object pattern))
                {
                    ((ValuePattern)pattern).SetValue(rdcIP);
                    Console.WriteLine("Window:{0} write successful!!", rdckey);

                }
                else
                {
                    TextboxElement.SetFocus();
                    SendKeys.SendWait("^{HOME}");   // Move to start of control
                    SendKeys.SendWait("^+{END}");   // Select everything
                    SendKeys.SendWait("{DEL}");     // Delete selection
                    SendKeys.SendWait("Something!");

                    // OR 
                    // SendMessage(element.Current.NativeWindowHandle, WM_SETTEXT, 0, "Something!");
                }
            }

            else
            {
                Console.WriteLine("Window:{0} write fail!!", rdckey);
            }
        }
        catch (ElementNotAvailableException)
        {

        }
    }

    public static void clickFButton(String y)
    {

        var notepad = System.Diagnostics.Process.GetProcessesByName("mstsc").FirstOrDefault();
        if (notepad!=null)
        {
            var root = AutomationElement.FromHandle(notepad.MainWindowHandle);
            var element = root.FindAll(TreeScope.Subtree, Condition.TrueCondition)
                                .Cast<AutomationElement>()
                                .Where(x => x.Current.ClassName=="Button"&&
                                x.Current.Name==y
                                            ).FirstOrDefault();
            if (element!=null)
            {
                InvokePattern invokePattern = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                invokePattern?.Invoke();
            }
            else
            {

            }

        }
    }

    //read config from notepad
    public static Dictionary<string, string> getConfig()
    {
        Dictionary<string, string> dictionary = new Dictionary<string, string>();
        using (var reader = new StreamReader("windowList.txt"))
        {
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                if (line==null) continue;
                var values = line.Split(';');
                dictionary.Add(values[0], values[1]);
            }
        }
        return dictionary;
    }


}