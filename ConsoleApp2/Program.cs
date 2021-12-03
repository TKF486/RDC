using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;

class Program
{
    public static Dictionary<string, string> dictionary;

    // Find Window 
    [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
    static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);

    // Activate an application window.
    [DllImport("USER32.DLL")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    static void Main(string[] args)
    {
        Logger("Program Start");
        dictionary=getConfig();
        while (true)
        {
            GetRDCWindows(dictionary);
            //set the timer for the program to rerun after a specific time
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
                AutomationElementCollection winCollection = rootElement.FindAll(TreeScope.Children, Condition.TrueCondition);

                int counter = 0;
                foreach (AutomationElement elementIter in winCollection)
                {
                    String elementName = elementIter.Current.Name;
                    if (elementName.Contains(rdc.Key))
                    {
                        counter++;
                        break;
                    }
                }
                if (counter==1)
                {
                    //RDC already exists
                }
                else
                {
                    Thread.Sleep(2000);
                    //open new rdc window
                    Process.Start("mstsc.exe");                    
                    Thread.Sleep(5000);
                    //take the desktop as root
                    AutomationElement rootElement2 = AutomationElement.RootElement;
                    //find all window in desktop
                    AutomationElementCollection winCollection2 = rootElement2.FindAll(TreeScope.Children, Condition.TrueCondition);
                    AutomationElement mstscWin = null;
                    foreach (AutomationElement elementIter2 in winCollection2)
                    {
                        String elementName = elementIter2.Current.Name;
                        if (elementName.Contains(windowsName))
                        {
                            mstscWin=elementIter2;
                            break;
                        }
                    }
                    var RDC_Win = mstscWin;
                    //locate the edit box in rdc
                    var RDC_EditBox = RDC_Win.FindAll(TreeScope.Subtree, Condition.TrueCondition).Cast<AutomationElement>().Where(x => (x.Current.AutomationId=="5012")).FirstOrDefault();
                    Thread.Sleep(5000);

                    if (RDC_EditBox!=null)
                    {
                        EnterValue(RDC_EditBox, rdc.Value, rdc.Key);
                        clickButton(RDC_Win, "Connect");
                        Thread.Sleep(2000);
                        clickButton(RDC_Win, "Yes");
                        Thread.Sleep(5000);
                        clickOKButton(rdc.Key);
                    }
                    Thread.Sleep(2000);

                }
                Thread.Sleep(5000);
            }
        }
        catch (Exception ex)
        {
        }
    }

    //enter ip address into rdc
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
                    //Console.WriteLine("Window:{0} write successful!!", rdckey);
                }
                else
                {
                    TextboxElement.SetFocus();
                }
            }
        }
        catch (ElementNotAvailableException e)
        {
        }
    }

    //Click all the buttons that is invlove to open rdc
    public static void clickButton(AutomationElement rDCElement1, String y)
    {
        try
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
                //button cannot be detected
            }
        }
        catch (ElementNotAvailableException e)
        {
        }
    }

    //pressing the ok button after rdc is opend
    public static void clickOKButton(string key)
    {
        try
        {
            var hWnd = FindWindowByCaption(IntPtr.Zero, key);
            if (hWnd==IntPtr.Zero)
            {
                //If cannot find window
            }
            else
            {
                //Sucess find open rdc
                SetForegroundWindow(hWnd);
                SendKeys.SendWait("{ENTER}");
            }
        }

        catch (Exception e)
        {
        }

    }

    //read config from notepad
    public static Dictionary<string, string> getConfig()
    {
        Dictionary<string, string> dictionary = new Dictionary<string, string>();
        try
        {
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
        }
        catch (Exception e)
        {
        }
        return dictionary;
    }

    public static void VerifyDir(string path)
    {
        try
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            if (!dir.Exists)
            {
                dir.Create();
            }
        }
        catch { }
    }

    public static void Logger(string lines)
    {
        string path = "C:/Log/";
        VerifyDir(path);
        string fileName = DateTime.Now.Day.ToString()+DateTime.Now.Month.ToString()+DateTime.Now.Year.ToString()+"_Logs.txt";
        try
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter(path+fileName, true);
            file.WriteLine(DateTime.Now.ToString()+": "+lines);
            file.Close();
        }
        catch (Exception) { }
    }
}