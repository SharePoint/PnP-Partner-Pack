using System;
using System.Diagnostics;

namespace OfficeDevPnP.PartnerPack.Infrastructure.Diagnostics
{
    /// <summary>
    /// Class to filter entries which contains some given words on Trace Listener
    /// </summary>
    public class RemoveWordsFilter : TraceFilter
    {

        public RemoveWordsFilter(String[] words)
        {
            _words = words;
        }

        public String[] _words { get; set; }

        override public bool ShouldTrace(TraceEventCache cache, string source,
            TraceEventType eventType, int id, string formatOrMessage,
            object[] args, object data, object[] dataArray)
        {
            if (formatOrMessage == null) return false;
            foreach (string word in _words)
            {
                if (formatOrMessage.Contains(word))
                {
                    return false;
                }
            }
            return true;
        }
    }

}
