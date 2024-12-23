using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class StateManager
    {
        private  static string fileStateName = "state.txt";

        public static void SaveApplicationState(int processedTickets)
        {
            File.WriteAllText(fileStateName, processedTickets.ToString());
        }

        public static int ReadApplicationState()
        {
            int result = -1;
            if(File.Exists(fileStateName))
            {
                var ticketCountStr = File.ReadAllText(fileStateName);
                int.TryParse(ticketCountStr, out result);
            }

            return result;
        }
    }
}
