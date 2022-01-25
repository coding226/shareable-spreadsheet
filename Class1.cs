using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Simulator
{
    class Class1
    {
        private static Mutex mutex = new Mutex();
        public void Example()
        {
            //Create mumber of thread to explain muiltiple thread example  
            SharableSpreadsheet sp = new SharableSpreadsheet(50, 560);
            for (int i = 0; i < 4; i++)
            {
                Thread t = new Thread(() => sp.getCell(500, 4));
                t.Name = string.Format("Thread {0} :", i + 1);
                t.Start();
            }
        }
       /* static void Main1(string[] args)
        {
            Class1 p = new Class1();
            p.Example();
        }*/
        //Method to implement syncronization using Mutex  
        static void MutexDemo(Object s, Object s1)
        {
            try
            {
                Console.WriteLine((String)s1);
                //Blocks the current thread until the current WaitHandle receives a signal.   
                mutex.WaitOne();   // Wait until it is safe to enter.  
                Console.WriteLine("{0} has entered in the Domain", Thread.CurrentThread.Name);
                Thread.Sleep((int)s);    // Wait until it is safe to enter.  
                Console.WriteLine("{0} is leaving the Domain\r\n", Thread.CurrentThread.Name);
            }
            finally
            {
                //ReleaseMutex unblock other threads that are trying to gain ownership of the mutex.  
                mutex.ReleaseMutex();
            }
        }
    }
}
