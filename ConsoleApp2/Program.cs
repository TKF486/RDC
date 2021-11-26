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
    //public static Process[] processList;

    static void Main(string[] args)
    {
        dictionary=getConfig();
        //processList=Process.GetProcesses();
        while (true)
        {
            GetRDCWindows(dictionary);
            Thread.Sleep(3000); 
        }     

    }


    //get RDC windows
    public static void GetRDCWindows(Dictionary<string, string> dictionary)
    {
        try
        {
            foreach (KeyValuePair<string, string> rdc in dictionary)
            {
                Process[] processlist = Process.GetProcesses();
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
                    Process.Start("mstsc.exe");
                    Thread.Sleep(5000);
                    var RDC = Process.GetProcessesByName("mstsc").FirstOrDefault();
                    Thread.Sleep(5000);
                    var RDCElement1 = AutomationElement.FromHandle(RDC.MainWindowHandle);
                    if(RDCElement1 == null)
                    {
                        Console.WriteLine("RDCElement1 is null!!");
                    }
                    var RDCElement1a = RDCElement1.FindAll(TreeScope.Subtree, Condition.TrueCondition);
                    if (RDCElement1a==null)
                    {
                        Console.WriteLine("RDCElement1a is null!!");
                    }
                    var RDCElement1b = RDCElement1a.Cast<AutomationElement>();
                    if (RDCElement1b==null)
                    {
                        Console.WriteLine("RDCElementb1 is null!!");
                    }
                    //Thread.Sleep(5000);
                    var RDCElement1c = RDCElement1b.Where(x => (x.Current.AutomationId=="5012"));
                    if (RDCElement1c==null)
                    {
                        Console.WriteLine("RDCElement1c is null!!");
                    }                   
                    var RDCElement1d = RDCElement1c.FirstOrDefault();
                    if (RDCElement1d==null)
                    {
                        Console.WriteLine("RDCElement1d is null!!");
                    }


                    Thread.Sleep(5000);


                    if (RDCElement1d!=null)
                    {
                        Console.WriteLine("textbox exists = ");
                        EnterValue(RDCElement1d, rdc.Value, rdc.Key);
                        clickFButton("Connect");
                        Thread.Sleep(2000);
                        clickFButton("Yes");
                    }
                    else
                    {
                        Console.WriteLine("text not exists");
                    }

                    Thread.Sleep(2000);

                }
                Thread.Sleep(5000);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
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
        catch (ElementNotAvailableException e)
        {
            Console.WriteLine(e);
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