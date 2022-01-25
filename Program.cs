using System;
using System.Collections.Generic;
using System.Threading;


namespace Simulator
{
    class Program
    {

        static int rows = 10;
        static int cols = 5;
        static int nThreads = 50;
        static int nOperations = 10;
        SharableSpreadsheet sheet;

        //Write Mutexes
        private static Mutex setCellMutex = new Mutex();
        private static Mutex writeMutex = new Mutex();
        private static Mutex setConcurrentSearchLimitMutex = new Mutex();

        //Read Semaphores
        private static Semaphore getCellSemaphore = new Semaphore(nThreads, nThreads);
        private static Semaphore searchStringSemaphore = new Semaphore(nThreads, nThreads);
        private static Semaphore searchInRowSemaphore = new Semaphore(nThreads, nThreads);
        private static Semaphore searchInColSemaphore = new Semaphore(nThreads, nThreads);
        private static Semaphore searchInRangeSemaphore = new Semaphore(nThreads, nThreads);
        private static Semaphore getSizeSemaphore = new Semaphore(nThreads, nThreads);
        private static Semaphore saveFileSemaphore = new Semaphore(1, nThreads);

        private static int semaphoreCount = nThreads;

        private static String filePath = "D:\\TamilVanan\\Temp\\";

        List<Thread> userThreads = new List<Thread>();
        List<String> randomWords = new List<String>();

        public Program(SharableSpreadsheet sheet)
        {
            this.sheet = sheet;
        }

        public static void Main(String[] args)
        {
            rows = Int16.Parse(args[0]);
            cols = Int16.Parse(args[1]);
            nThreads = Int16.Parse(args[2]);
            nOperations = Int16.Parse(args[3]);

            semaphoreCount = nThreads;

            //create a rows* cols spreadsheet
            SharableSpreadsheet sheet = new SharableSpreadsheet(rows, cols);

            //initialize values in Sheet
            InitializeValuesInSheet(sheet);

            //Weave Thread
            Program program = new Program(sheet);

            program.WeaveThread(nThreads);

            foreach (Thread userThread in program.userThreads)
            {
                userThread.Join();
            }

            sheet.save(filePath + "Final_Excel.csv");
            program.printSheet();
        }

        private void printSheet()
        {
            string[,] sheetData = this.sheet.getSheet();

            for (int i = 0; i < sheetData.GetLength(0); i++)
            {
                for (int j = 0; j < sheetData.GetLength(1); j++)
                {
                    String value = sheetData[i, j];
                    Console.Write(value + "\t");
                }
                Console.WriteLine("");
            }
        }

        private void WeaveThread(int nThreads)
        {
            for(int i = 0; i < nThreads; i++)
            {
                Thread t = new Thread(() => dOperations(nOperations));
                t.Name = string.Format("User [{0}] ", t.ManagedThreadId);
                userThreads.Add(t);
                t.Start();
            }
        }

        private void dOperations(int nOperations)
        {
            Dictionary<string, Func<int>> writeFunctions = new Dictionary<string, Func<int>>();
            writeFunctions["setCell"] = this.setCell;
            writeFunctions["write"] = this.write;

            Dictionary<string, Func<int>> readFunctions = new Dictionary<string, Func<int>>();
            readFunctions["getCell"] = this.getCell;
            readFunctions["searchString"] = this.searchString;
            readFunctions["getSize"] = this.getSize;
            readFunctions["searchInRow"] = this.searchInRow;
            readFunctions["searchInCol"] = this.searchInCol;
            readFunctions["searchInRange"] = this.searchInRange;
            readFunctions["setConcurrentSearchLimit"] = this.setConcurrentSearchLimit;
            readFunctions["saveFile"] = this.saveFile;


            for (int i = 0; i < nOperations; i++)
            {
                int readOrWrite = GetRandomNumberBetween(1, 10);
                if(readOrWrite <= 5)
                {
                    //Read Operations using Semaphores
                    Func<int> randomMethod = getRandomMethod(readFunctions);
                    randomMethod();
                } else
                {
                    //Write Operations using Mutex
                    Func<int> randomMethod = getRandomMethod(writeFunctions);
                    randomMethod();
                }
                Thread.Sleep(100);
            }
        }

        private int write()
        {
            writeMutex.WaitOne();
            Dictionary<string, Func<int>> writeFunctions = new Dictionary<string, Func<int>>();
            writeFunctions["exchangeRows"] = this.exchangeRows;
            writeFunctions["exchangeCols"] = this.exchangeCols;
            writeFunctions["addRow"] = this.addRow;
            writeFunctions["addCol"] = this.addCol;

            Func<int> randomMethod = getRandomMethod(writeFunctions);
            randomMethod();
            writeMutex.ReleaseMutex();
            return 0;
        }

        private int saveFile()
        {
            saveFileSemaphore.WaitOne();

            String threadName = Thread.CurrentThread.Name;
            String fileName = threadName + ".csv";
            this.sheet.save(filePath + fileName);

            Console.WriteLine(threadName + " saved File " + fileName);

            saveFileSemaphore.Release();
            return 0;
        }

        private int setConcurrentSearchLimit()
        {
            setConcurrentSearchLimitMutex.WaitOne();
            int count = GetRandomNumberBetween(1, nThreads);
            String threadName = Thread.CurrentThread.Name;

            this.sheet.setConcurrentSearchLimit(count);
            
            Console.WriteLine(threadName + " Search ConcurrentSearchLimit set to " + count);
            setConcurrentSearchLimitMutex.ReleaseMutex();
            return 0;
        }

        private Func<int> getRandomMethod(Dictionary<string, Func<int>> writeFunctions)
        {
            int count = writeFunctions.Count;
            int randNum = GetRandomNumberBetween(0, count);

            int loopCount = 0;
            foreach (KeyValuePair<string, Func<int>> ele in writeFunctions)
            {
                if(loopCount == randNum)
                {
                    return ele.Value;
                }
                loopCount++;
            }
            return null;
        }

        private int searchInRange()
        {
            searchInRangeSemaphore.WaitOne();
            int row1 = GetRandomNumberBetween(0, this.sheet.getSheet().GetLength(0) - 1);
            int col1 = GetRandomNumberBetween(0, this.sheet.getSheet().GetLength(1) - 1);

            int row2 = GetRandomNumberBetween(row1, this.sheet.getSheet().GetLength(0) - 1);
            int col2 = GetRandomNumberBetween(col1, this.sheet.getSheet().GetLength(1) - 1);

            int searchRow = 0;
            int searchCol = 0;

            String word = RetreiveARandomWords();
            String threadName = Thread.CurrentThread.Name;

            bool result = this.sheet.searchInRange(col1, col2, row1, row2, word, ref searchRow, ref searchCol);

            if(result)
            {
                Console.WriteLine(threadName + " Search String " + word + " in range: row1: " + row1 + " col1: " + col1 + " row2: " + row2 + " col2: " + col2  + ". Found in row: " + searchRow + "col: " + searchCol);
            } else
            {
                Console.WriteLine(threadName + " Search String " + word + " in range: row1: " + row1 + " col1: " + col1 + " row2: " + row2 + " col2: " + col2 + ". String not found in Range");
            }

            searchInRangeSemaphore.Release();
            return 0;
        }

        private int searchInCol()
        {
            searchInColSemaphore.WaitOne();
            int col = GetRandomNumberBetween(0, this.sheet.getSheet().GetLength(1) - 1);
            int searchRow = 0;

            String word = RetreiveARandomWords();
            String threadName = Thread.CurrentThread.Name;

            bool result = this.sheet.searchInCol(col, word, ref searchRow);

            if(result)
            {
                Console.WriteLine(threadName + " Search String " + word + " in col: " + col + ". Found in row: " + searchRow);
            } else
            {
                Console.WriteLine(threadName + " Search String " + word + " in col: " + col + ". Not Found");
            }
            
            searchInColSemaphore.Release();
            return 0;
        }

        private int searchInRow()
        {
            searchInRowSemaphore.WaitOne();
            int row = GetRandomNumberBetween(0, this.sheet.getSheet().GetLength(0) - 1);
            int searchCol = 0;

            String word = RetreiveARandomWords();
            String threadName = Thread.CurrentThread.Name;

            bool result = this.sheet.searchInRow(row, word, ref searchCol);

            if(result)
            {
                Console.WriteLine(threadName + " Search String " + word + " in row: " + row + ". Found in col: " + searchCol);
            } else
            {
                Console.WriteLine(threadName + " Search String " + word + " in row: " + row + ". Not Found");
            }

            searchInRowSemaphore.Release();
            return 0;
        }

        private int getSize()
        {
            getSizeSemaphore.WaitOne();
            int row = 0;
            int col = 0;
            String threadName = Thread.CurrentThread.Name;
            this.sheet.getSize(ref row, ref col);

            Console.WriteLine(threadName + " Size of Sheet is row: " + row + " col: " + col);
            getSizeSemaphore.Release();
            return 0;
        }

        private int searchString()
        {
            searchStringSemaphore.WaitOne();

            int row = 0;
            int col = 0;
            String word = RetreiveARandomWords();
            String threadName = Thread.CurrentThread.Name;
            this.sheet.searchString(word, ref row, ref col);

            Console.WriteLine(threadName + " Search String " + word + " is present in row: " + row + " col: " + col);
            searchStringSemaphore.Release();
            
            return 0;
        }

        private int getCell()
        {
            getCellSemaphore.WaitOne();

            int row = GetRandomNumberBetween(0, rows - 1);
            int col = GetRandomNumberBetween(0, cols - 1);
            String value = this.sheet.getCell(row, col);
            String threadName = Thread.CurrentThread.Name;

            Console.WriteLine(threadName + " Value in cell " + row + " " + col + " is " + value);

            getCellSemaphore.Release();
            return 0;
        }

        private int setCell()
        {
            setCellMutex.WaitOne();
            int row = GetRandomNumberBetween(0, rows - 1);
            int col = GetRandomNumberBetween(0, cols - 1);
            String word = GenerateRandomWord(GetRandomNumberBetween(1, 10));
            String threadName = Thread.CurrentThread.Name;

            this.sheet.setCell(row, col, word);
            
            Console.WriteLine(threadName + " String " + word + " set in cell " + row + " " + col);

            setCellMutex.ReleaseMutex();
            return 0;
        }

        private int addRow()
        {
            int row = GetRandomNumberBetween(1, rows);
            rows += 1;
            String threadName = Thread.CurrentThread.Name;

            this.sheet.addRow1(row);

            Console.WriteLine(threadName + " Added a new row after row " + row);

            return 0;
        }

        private int addCol()
        {
            int col = GetRandomNumberBetween(1, cols);
            cols += 1;
            String threadName = Thread.CurrentThread.Name;

            this.sheet.addCol(col);

            Console.WriteLine(threadName + " Added a new column after column " + col);

            return 0;
        }

        private int exchangeRows()
        {
            int row1 = GetRandomNumberBetween(1, rows);
            int row2 = GetRandomNumberBetween(1, rows);
            String threadName = Thread.CurrentThread.Name;

            this.sheet.exchangeRows(row1, row2);

            Console.WriteLine(threadName + " Exchanged the rows " + row1 + " and " + row2);

            return 0;
        }

        private int exchangeCols()
        {
            int col1 = GetRandomNumberBetween(0, cols - 1);
            int col2 = GetRandomNumberBetween(0, cols - 1);
            String threadName = Thread.CurrentThread.Name;

            this.sheet.exchangeCols(col1, col2);

            Console.WriteLine(threadName + " Exchanged the cols " + col1 + " and " + col2);

            return 0;
        }

        private int GetRandomNumberBetween(int v1, int v2)
        {
            Random rand = new Random();
            int number = rand.Next(v1, v2);
            return number;
        }

        private static void InitializeValuesInSheet(SharableSpreadsheet sheet)
        {
            String[,] sheetValues = sheet.getSheet();
            for (int i = 0; i < rows; i++)
            {
                for(int j = 0; j < cols; j++)
                {
                    String value = "Cell_" + i + "_" + j;
                    sheetValues[i, j] = value;
                }
            }
        }

        private String GenerateRandomWord(int len)
        {
            char[] letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
            Random rand = new Random();
            // Make a word.
            string word = "";
            for (int j = 1; j <= len; j++)
            {
                // Pick a random number between 0 and 25
                // to select a letter from the letters array.
                int letter_num = rand.Next(0, letters.Length - 1);

                // Append the letter.
                word += letters[letter_num];
            }
            randomWords.Add(word);
            return word;
        }

        private String RetreiveARandomWords()
        {
            int count = randomWords.Count;
            String word = null;
            if (count > 0)
            {
                int rand = GetRandomNumberBetween(0, count);
                word = randomWords[rand];
            } else
            {
                word = GenerateRandomWord(10);
            }
            
            return word;
        }

        public static bool UpdateConcurrentSearchThreadCount(int threadCount)
        {
            if (semaphoreCount > threadCount)
            {
                int waitCount = semaphoreCount - threadCount;
                for (int i = 0; i < waitCount; i++)
                {
                    searchStringSemaphore.WaitOne();
                    searchInRowSemaphore.WaitOne();
                    searchInColSemaphore.WaitOne();
                    searchInRangeSemaphore.WaitOne();
                }
                
                semaphoreCount = threadCount;
            } else
            {
                int releaseCount = threadCount - semaphoreCount;
                if(releaseCount > 0)
                {
                    searchStringSemaphore.Release(releaseCount);
                    searchInRowSemaphore.Release(releaseCount);
                    searchInColSemaphore.Release(releaseCount);
                    searchInRangeSemaphore.Release(releaseCount);
                    semaphoreCount = threadCount;
                }
            }
            return true;
        }
    }
}