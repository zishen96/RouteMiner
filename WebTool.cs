using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Diagnostics;

namespace RouteMiner
{
    public class WebTool
    {
        //Base URL for USPS Address validation API
        private const string BaseURL = "http://production.shippingapis.com/ShippingAPI.dll";
        private WebClient wsClient = new WebClient();

        //Replace with USPS login
        private string USPS_UserID = "1234ABCD";
        private string USPS_Password = "1234ABCD";

        public WebTool()
        {

        }

        //Constructor with User ID parameter/Password if needed
        public WebTool(string NewUserID, string NewPassword)
        {
            USPS_UserID = NewUserID;
            USPS_Password = NewPassword;
        }

        private string GetDataFromSite(string USPS_Request)
        {
            string strResponse = "";

            //Send the request to USPS
            byte[] ResponseData = wsClient.DownloadData(USPS_Request);

            //Convert byte stream to string data
            foreach (byte item in ResponseData)
            {
                strResponse += (char)item;
            }
            return strResponse;
        }

        //Method providing interface to USPS Address Validation API
        public string AddressValidateRequest(string Address1, string Address2, string City, string State, string Zip5, string Zip4)
        {
            string strResponse = "", strUSPS = "";

            #region API usage example
            /*
            http://production.shippingapis.com/ShippingAPI.dll?API=Verify&XML=
            <AddressValidateRequest USERID="1234ABCD" PASSWORD="1234ABCD">
            <Revision>1</Revision>
            <Address ID="0">
            <Address1></Address1>
            <Address2>4901 Evergreen Road</Address2>
            <City>Dearborn</City>
            <State>MI</State>
            <Zip5>48128</Zip5>
            <Zip4></Zip4>
            </Address>
            </AddressValidateRequest>
            */
            #endregion

            strUSPS = BaseURL + "?API=Verify&XML=<AddressValidateRequest USERID=\"" + USPS_UserID + "\" PASSWORD=\"" + USPS_Password + "\">";
            strUSPS += "<Revision>1</Revision>";
            strUSPS += "<Address ID=\"0\">";
            strUSPS += "<Address1>" + Address1 + "</Address1>";
            strUSPS += "<Address2>" + Address2 + "</Address2>";
            strUSPS += "<City>" + City + "</City>";
            strUSPS += "<State>" + State + "</State>";
            strUSPS += "<Zip5>" + Zip5 + "</Zip5>";
            strUSPS += "<Zip4>" + Zip4 + "</Zip4>";
            strUSPS += "</Address></AddressValidateRequest>";

            Debug.WriteLine("XML Request [START]");
            Debug.WriteLine(strUSPS);
            Debug.WriteLine("XML Request [END]");

            //Send the request to USPS
            Debug.WriteLine("XML Response [START]");
            strResponse = GetDataFromSite(strUSPS);
            Debug.WriteLine(strResponse);
            Debug.WriteLine("XML Response [END]");

            return strResponse;
        }

        //Method for USPS API
        public string AddressValidateReq(BuilderProduct product)
        {
            string strResponse = "";
            string strUSPS = "";

            strUSPS = BaseURL + "?API=Verify&XML=<AddressValidateRequest USERID=\"" + USPS_UserID + "\" PASSWORD=\"" + USPS_Password + "\">";
            strUSPS += "<Revision>1</Revision>";
            strUSPS += "<Address ID=\"0\">";
            strUSPS += "<Address1>" + "</Address1>";
            strUSPS += "<Address2>" + $"{product.StreetNum} {product.StreetName} {product.AptNum} " + "</Address2>";
            strUSPS += "<City>" + $"{product.City} " + "</City>";
            strUSPS += "<State>" + $"{product.State} " + "</State>";
            strUSPS += "<Zip5>" + $"{product.Zip} " + "</Zip5>";
            strUSPS += "<Zip4>" + "</Zip4>";
            strUSPS += "</Address></AddressValidateRequest>";

            //Send request to USPS
            strResponse = GetDataFromSite(strUSPS);

            return strResponse;
        }
    }
}
