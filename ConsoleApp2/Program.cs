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
    static void Main(string[] args)
    {
        Dictionary<string, string> dictionary = getConfig();
        Process[] processlist = Process.GetProcesses();
        GetRDCWindows(dictionary, processlist);
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
                                        (x.Current.Name=="Connect") ||(x.Current.Name=="Yes")).FirstOrDefault();

                    EnterValue(RDCElement1, rdc.Value, rdc.Key);
                    
                }
                Thread.Sleep(5000);
            }


            Thread.Sleep(10000);
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
        }
        catch (ElementNotAvailableException)
        {

        }
    }

    private static void ClickConnectButton(AutomationElement RDCElement1, String y)
    {
        try
        {
            AutomationElement ButtonElement = RDCElement1.FindAll(TreeScope.Subtree, Condition.TrueCondition)
                                            .Cast<AutomationElement>()
                                            .Where(x => x.Current.ClassName=="Button"&&x.Current.Name=="y")
                                            .FirstOrDefault();

            InvokePattern invokePattern = ButtonElement.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern?.Invoke();
        }
        catch (ElementNotAvailableException)
        {

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