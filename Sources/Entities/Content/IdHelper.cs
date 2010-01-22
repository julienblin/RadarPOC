using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Russell.RADAR.POC.Entities.Content
{
    public static class IdHelper
    {
        public static Random random = new Random();

        public static string GenerateRandomId()
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < 8; i++)
            {
                //26 letters in the alfabet, ascii + 65 for the capital letters
                builder.Append(Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65))));
            }
            return builder.ToString();
        }

        public static int GenerateIntId()
        {
            return (random.Next(900) + 100);
        }
    }
}
