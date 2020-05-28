using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RouteMiner
{
    public class ParsedResponse
    {
        public List<string[]> AddressPlusCarrier;
        public Dictionary<string, int> CarrierPlusTally;

        public ParsedResponse()
        {
            AddressPlusCarrier = new List<string[]>();
            CarrierPlusTally = new Dictionary<string, int>();
        }
    }

    public class ResponseA
    {
        string _address;
        string _carrier;

        public ResponseA(string address, string carrier)
        {
            _address = address;
            _carrier = carrier;
        }

        public string Address { get { return _address; } }
        public string Carrier { get { return _carrier; } }
    }

    public class ResponseB
    {
        string _carrier;
        int _count;

        public ResponseB(string carrier, int val)
        {
            _carrier = carrier;
            _count = val;
        }

        public string Carrier { get { return _carrier; } }
        public int Count { get { return _count; } }
    }
}
