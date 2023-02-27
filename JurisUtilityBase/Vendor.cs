using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JurisUtilityBase
{
    public class Vendor
    {

        public Vendor()
        {
            num = "";
            name = "";
            addy = "";
            city = "";
            state = "";
            zip = "";
            fein = "";
            hasFEIN = false;
            hasError = false;
            errorMessage = "";
        }

        public string num { get; set; }
        public string name { get; set; }
        public string addy { get; set; }
        public string city { get; set; }
        public string state { get; set; }
        public string zip { get; set; }
        public string fein { get; set; }
        public bool hasFEIN { get; set; }

        public bool hasError { get; set; }

        public string errorMessage { get; set; }
    }
}
