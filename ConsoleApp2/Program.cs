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
                String windowsName = "Remote Desktop Connection";
                AutomationElement rootElement = AutomationElement.RootElement;
                AutomationElement element = null;
                AutomationElementCollection winCollection = rootElement.FindAll(TreeScope.Children, Condition.TrueCondition);

                Console.WriteLine("Checking for:{0} ...", rdc.Key);
                int counter = 0;
                foreach (AutomationElement elementIter in winCollection)
                {
                    String elementName = elementIter.Current.Name;
                    if (elementName.Contains(rdc.Key))
                    {
                        counter++;
                        break;
                    }
                    else
                    {

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
                    AutomationElement rootElement2 = AutomationElement.RootElement;
                    AutomationElementCollection winCollection2 = rootElement2.FindAll(TreeScope.Children, Condition.TrueCondition);
                    AutomationElement mstscWin = null;
                    foreach (AutomationElement elementIter2 in winCollection2)
                    {
                        String elementName = elementIter2.Current.Name;
                        //Console.WriteLine(elementName);
                        if (elementName.Contains(windowsName))
                        {
                            mstscWin=elementIter2;
                            Console.WriteLine("Success find mstsc!!");
                            break;

                        }

                        else
                        {

                        }

                    }
                    //var RDC = Process.GetProcessesByName("mstsc").FirstOrDefault();
                    //Thread.Sleep(5000);
                    var RDCElement1 = mstscWin;
                    if (RDCElement1==null)
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
                        clickFButton(RDCElement1, "Connect");
                        Thread.Sleep(2000);
                        clickFButton(RDCElement1, "Yes");
                        Thread.Sleep(5000);
                        clickOKButton(rdc.Key);
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

    public static void clickFButton(AutomationElement rDCElement1, String y)
    {
        var element = rDCElement1.FindAll(TreeScope.Subtree, Condition.TrueCondition)
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
            Console.WriteLine("Cannot find button!!");
        }
    }

    public static void clickOKButton(string key)
    {
        AutomationElement rootElement = AutomationElement.RootElement;
        AutomationElementCollection winCollection2 = rootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        AutomationElement targetWin = null;
        foreach (AutomationElement elementIter in winCollection2)
        {
            String elementName = elementIter.Current.Name;
            //Console.WriteLine(elementName);
            if (elementName.Contains(key))
            {
                Console.WriteLine();
                Console.WriteLine("Success find the opend rdc: "+key);
                targetWin=elementIter;
                break;
            }
            else
            {
                //Console.WriteLine("fail to find the opend rdc!!");
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