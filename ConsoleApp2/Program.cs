﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Automation;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.IO;

namespace Test
{
    class progam
    {
        static void Main(string[] args)
        {
           
                Semaphore gate = new Semaphore(2,2);
                Semaphore writeGate = new Semaphore(1,1);
                
                Dictionary<string, string> dictionary = getConfig();
                //continue loop through the process every 5s
                while (true)
                {
                    //program alive counter, check if the program is open
                    //int counter = 0;
                    //get all processes
                    Process[] processlist = Process.GetProcesses();

                
                Parallel.ForEach(dictionary, rdc =>
                {

                        //gate.WaitOne();
         
                        Console.WriteLine("Checking for:{0} ...", rdc.Key);
                        int counter = 0;
                        foreach (Process process in processlist)

                        {
                           
                            if (!String.IsNullOrEmpty(process.MainWindowTitle))
                            {
                                //Console.WriteLine("ID: {0} = Name: {1}", process.Id, process.MainWindowTitle);

                                if (process.MainWindowTitle == rdc.Key)
                                {
                                    counter++;
                                    //Console.WriteLine("Window:{0} is not opened!!", rdc.Key);  
                                }
                                else
                                {
                                    //Console.WriteLine("no found!!");
                                }



                                //check if RDC is opened


                            }

                        }

                        if (counter == 1)
                        {
                            Console.WriteLine("Window:{0} is already opened!!", rdc.Key);
                        }
                        else
                        {
                            Thread.Sleep(2000);
                            Console.WriteLine("Window:{0} is not opened!!", rdc.Key);
                            startProgram();
                            Thread.Sleep(2000);
                        gate.WaitOne();

                        writeLine(rdc.Value, ref writeGate, ref gate, rdc.Key);
                            Thread.Sleep(5000);
                            clickFButton("Connect", rdc.Key);
                            Thread.Sleep(2000);
                            clickFButton("Yes", rdc.Key);
                        }

                        gate.Release();
                    Thread.Sleep(3000);
                });



                //sleep 5s
                Thread.Sleep(3000);

                }

                /////////////////////////////////
            }






            //start the RDC prgram
            public static void startProgram()
        {
            Process firstApp = new Process();

            try
            {
                firstApp.StartInfo.FileName = "mstsc.exe";
                firstApp.Start();
                //Console.WriteLine("ID: {0} = Name: {1}", firstApp.Id, firstApp.ProcessName);
                //firstApp.WaitForExit();


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public static void writeLine(String rdcIP, ref Semaphore writeGate, ref Semaphore gate, string key)
        {
            try
            {
                //gate.WaitOne();
                writeGate.WaitOne();
                //Console.WriteLine(writeGate);
               
                var notepad = System.Diagnostics.Process.GetProcessesByName("mstsc").FirstOrDefault();
                if (notepad != null)
                {
                    var root = AutomationElement.FromHandle(notepad.MainWindowHandle);
                    var element = root.FindAll(TreeScope.Subtree, Condition.TrueCondition)
                                        .Cast<AutomationElement>()
                                        .Where(x => x.Current.ClassName == "Edit"
                                                    ).FirstOrDefault();
                    if (element != null)
                    {
                        if (element.TryGetCurrentPattern(ValuePattern.Pattern, out object pattern))
                        {
                            ((ValuePattern)pattern).SetValue(rdcIP);
                            Console.WriteLine("Program{0}: WriteLine succeed!!", key);
                            writeGate.Release();
                        }
                        else
                        {
                            Console.WriteLine("Cannot found!!");
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine("Program{0}: WriteLine fail", key);
            }
        }

        public static void clickFButton(String y, string key)
        {
            try
            {
                var notepad = System.Diagnostics.Process.GetProcessesByName("mstsc").FirstOrDefault();
                if (notepad != null)
                {
                    var root = AutomationElement.FromHandle(notepad.MainWindowHandle);
                    var element = root.FindAll(TreeScope.Subtree, Condition.TrueCondition)
                                        .Cast<AutomationElement>()
                                        .Where(x => x.Current.ClassName == "Button" &&
                                        x.Current.Name == y
                                                    ).FirstOrDefault();
                    if (element != null)
                    {
                        InvokePattern invokePattern = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        invokePattern?.Invoke();
                        Console.WriteLine("Program{0}: Press Button succeed!!", key);
                    }
                    else
                    {

                    }

                }
            }

            catch (Exception ex)
            {
                Console.WriteLine("Program{0}: Press button fail", key);
            }





        }

  

        public static Dictionary<string, string> getConfig()
        {
            Dictionary<string, string> dictionary = new Dictionary<string, string>();
            using (var reader = new StreamReader("windowList.txt"))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    if (line == null) continue;
                    var values = line.Split(';');
                    dictionary.Add(values[0], values[1]);
                }
            }

            return dictionary;

        }

    }
}