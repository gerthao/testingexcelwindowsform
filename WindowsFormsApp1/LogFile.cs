using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class LogFile
    {
        private string name;
        private string text;
        private int lineCount;
        private TextWriter writer;
        public LogFile() : this("Log.txt")
        {
        }
        public LogFile(string name) {
            this.name = name;
            writer = new StreamWriter(this.name);
            lineCount = 0;
        }
        public void Log(string message)
        {
            writer.WriteLine($"[{""+lineCount++}, {DateTime.Now.ToLongTimeString()}]: {message}");
        }
        public void SetWriter(TextWriter tw) => writer = tw;
        public void Close() => writer.Close();
    }
}
