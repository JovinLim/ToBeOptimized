using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace TBO_Plugin
{
    public enum RequestId : int
    {
        None = 0,
        Iteration_1 = 1,
        Iteration_2 = 2,
        Iteration_3 = 3,
        Iteration_4 = 4,
        Iteration_5 = 5,
        Iteration_6 = 6,
        Iteration_7 = 7,
        Iteration_8 = 8,
        Iteration_9 = 9,
        Iteration_10 = 10,
        Iteration_11 = 11,
        Iteration_12 = 12,
        Apply = 13,
        Cancel = 14,
        FireEgress = 15,
        HideFE = 16,
    }

    public class Request
    {
        // Storing the value as a plain Int makes using the interlocking mechanism simpler
        private int m_request = (int)RequestId.None;
        
        /// <summary>
        ///   Take - The Idling handler calls this to obtain the latest request. 
        /// </summary>
        /// <remarks>
        ///   This is not a getter! It takes the request and replaces it
        ///   with 'None' to indicate that the request has been "passed on".
        /// </remarks>
        /// 
        public RequestId Take()
        {
            return (RequestId)Interlocked.Exchange(ref m_request, (int)RequestId.None);
        }

        /// <summary>
        ///   Make - The Dialog calls this when the user presses a command button there. 
        /// </summary>
        /// <remarks>
        ///   It replaces any older request previously made.
        /// </remarks>
        /// 
        public void Make(RequestId request)
        {
            Interlocked.Exchange(ref m_request, (int)request);
        }
    }
}
