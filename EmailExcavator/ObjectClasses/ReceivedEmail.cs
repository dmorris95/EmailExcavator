using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailExcavator.ObjectClasses
{
    public class ReceivedEmail
    {
        public string Subject { get; set; }
        public string From { get; set; }
        public DateTime ReceivedDateTime { get; set; }
        
    }
}
