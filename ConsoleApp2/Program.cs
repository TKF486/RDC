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
                                        (x.Current.Name=="Connect")).FirstOrDefault();
                    AutomationElement RDCElement3 = AutomationElement.FromHandle(RDC.MainWindowHandle)
                                    .FindAll(TreeScope.Subtree, Condition.TrueCondition)
                                        .Cast<AutomationElement>()
                                        .Where(x => x.Current.ClassName=="Button"&&
                                      (x.Current.Name=="Yes")).FirstOrDefault();

                    Thread.Sleep(5000);
                    if (RDCElement1!=null)
                    {
                        Console.WriteLine("textbox exists = " );
                        EnterValue(RDCElement1, rdc.Value, rdc.Key);
                    }
                    else
                    {
                        Console.WriteLine("text not exists");
                    }
         
                    Thread.Sleep(2000);
                    if (RDCElement2!=null)
                    {
                        Console.WriteLine("connect btn exists = ");
                        Thread.Sleep(2000);
                        ClickConnectButton(RDCElement2, rdc.Key);
                    }
                    else
                    {
                        Console.WriteLine("connect btn not exists");
                    }

                   
            
                   


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

    private static void ClickConnectButton(AutomationElement RDCElement, string key)
    {
        try
        {
            AutomationElement ButtonElement = RDCElement.FindAll(TreeScope.Subtree, Condition.TrueCondition)
                                        .Cast<AutomationElement>()
                                        .Where(x => x.Current.ClassName=="Button"&&
                                        (x.Current.Name=="Connect")).FirstOrDefault();           
            if(ButtonElement !=null)
            {
                InvokePattern invokePattern = ButtonElement.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                invokePattern?.Invoke();
                Console.WriteLine("Window:{0} press Connect button success!!", key);
            }
            else
            {
                Console.WriteLine("Window:{0} press Connect button fail!!", key);
            }
            
        }
        catch (ElementNotAvailableException)
        {

        }
    }

    private static void ClickYesButton(AutomationElement RDCElement, string key)
    {
        try
        {
            AutomationElement ButtonElement = RDCElement.FindAll(TreeScope.Subtree, Condition.TrueCondition)
                                        .Cast<AutomationElement>()
                                        .Where(x => x.Current.ClassName=="Button"&&
                                        (x.Current.Name=="Yes")).FirstOrDefault();
            if (ButtonElement!=null)
            {
                InvokePattern invokePattern = ButtonElement.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                invokePattern?.Invoke();
                Console.WriteLine("Window:{0} press Yes button success!!", key);
            }
            else
            {
                Console.WriteLine("Window:{0} press Yes button fail!!", key);
            }

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