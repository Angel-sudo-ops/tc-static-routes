using System;
using TwinCAT.Ads;

class Program
{
    static void Main()
    {
        // Retrieve the local AMS Net ID using the ADS API
        AmsNetId localAmsNetId = AmsNetId.Local;
        Console.WriteLine("Local AMS Net ID: " + localAmsNetId.ToString());
    }
}