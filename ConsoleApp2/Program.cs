using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Threading;
using System.Timers;
using Timer = System.Timers.Timer;


class Program
{
    static int Main(string[] args)
    {
        Console.WriteLine("Enter \'q\' to quit the sample.");

        // Create a timer with a ten second interval.
        Timer aTimer = new Timer(10000);
        // Hook up the Elapsed event for the timer. 
        aTimer.Elapsed += StartThreads;
        aTimer.AutoReset = true;
        aTimer.Enabled = true;

        //Quit Sample
        while (Console.Read() == 'q')
        {
            return -1;
        }

        Console.ReadLine();
        aTimer.Stop();
        aTimer.Dispose();

        return 0;
    }

    public static void StartThreads(object o, ElapsedEventArgs e)
    {
        string RDCWindowsName = "Remote Desktop Connection";
        string NotepadWindowsName = "Untitled - Notepad";
        AutomationElement rootElement = AutomationElement.RootElement;
        var th = new Thread(() => GetRDCWindows(RDCWindowsName, rootElement));
        var th2 = new Thread(() => GetNotepadWindows(NotepadWindowsName, rootElement));
        th.Start();
        th2.Start();
    }

    private static void GetNotepadWindows(string NotepadWindowsName, AutomationElement rootElement)
    {
        try
        {
            //Retrieves a collection of all the immediate children of the desktop.
            AutomationElement NotepadElement = null;
            AutomationElementCollection winCollection = rootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
            bool NotepadOn = false;
            int iter = 1;

            foreach (AutomationElement element in winCollection)
            {
                String elementName = element.Current.Name;
                if (elementName.Contains(NotepadWindowsName))
                {
                    NotepadOn = true;
                    NotepadElement = element;
                }
                iter++;
            }

            //Open RDC if it is not opened
            if (NotepadOn == false)
            {
                System.Diagnostics.Process.Start("notepad.exe");
                Thread.Sleep(2000);
                var Notepad = System.Diagnostics.Process.GetProcessesByName("notepad").FirstOrDefault();
                NotepadElement = AutomationElement.FromHandle(Notepad.MainWindowHandle)
                                .FindAll(TreeScope.Subtree, Condition.TrueCondition)
                                .Cast<AutomationElement>()
                                .Where(x => x.Current.ClassName == "Edit")
                                .FirstOrDefault();
                EnterValue(NotepadElement);
            }
            else
            {
                EnterValue(NotepadElement);
            }
        }
        catch (ElementNotAvailableException)
        {

        }
    }


    //Function to find RDC Windows
    private static void GetRDCWindows(string RDCWindowsName, AutomationElement rootElement)
    {
        try
        {
            //Retrieves a collection of all the immediate children of the desktop.
            AutomationElement RDCElement = null;
            AutomationElementCollection winCollection = rootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
            bool RDCOn = false;
            int iter = 1;

            foreach (AutomationElement element in winCollection)
            {
                String elementName = element.Current.Name;
                if (elementName.Contains(RDCWindowsName))
                {
                    RDCOn = true;
                    RDCElement = element;
                }
                iter++;
            }

            //Open RDC if it is not opened
            if (RDCOn == false)
            {
                System.Diagnostics.Process.Start("mstsc.exe");
                Thread.Sleep(2000);
                var RDC = System.Diagnostics.Process.GetProcessesByName("mstsc").FirstOrDefault();
                RDCElement = AutomationElement.FromHandle(RDC.MainWindowHandle)
                                .FindAll(TreeScope.Subtree, Condition.TrueCondition)
                                .Cast<AutomationElement>()
                                .Where(x => x.Current.ClassName == "Edit")
                                .FirstOrDefault();
                EnterValue(RDCElement);
            }
            else
            {
                EnterValue(RDCElement);
            }
        }
        catch (ElementNotAvailableException)
        {

        }
    }

    private static void EnterValue(AutomationElement element)
    {
        try
        {
            AutomationElement TextboxElement = element.FindAll(TreeScope.Subtree, Condition.TrueCondition)
                                            .Cast<AutomationElement>()
                                            .Where(x => x.Current.ClassName == "Edit")
                                            .FirstOrDefault();
            if (TextboxElement != null)
            {
                if (TextboxElement.TryGetCurrentPattern(ValuePattern.Pattern, out object pattern))
                {
                    ((ValuePattern)pattern).SetValue("Something!");
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

    private static void ClickConnectButton(AutomationElement RDCElement)
    {
        try
        {
            AutomationElement ButtonElement = RDCElement.FindAll(TreeScope.Subtree, Condition.TrueCondition)
                                            .Cast<AutomationElement>()
                                            .Where(x => x.Current.ClassName == "Button" && x.Current.Name == "Connect")
                                            .FirstOrDefault();

            InvokePattern invokePattern = ButtonElement.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern?.Invoke();
        }
        catch (ElementNotAvailableException)
        {

        }
    }
}