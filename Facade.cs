using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;


namespace RouteMiner
{
    public class Facade
    {
        private Excel _excel;
        private WebTool _webTool;
        private Director _director;
        private Builder _builder;
        private BuilderProduct _bproduct;
        public ParsedResponse _parsedResponse;

        public Facade()
        {
            _webTool = new WebTool();
            _builder = new SBuilder();
            _director = new Director();
            _parsedResponse = new ParsedResponse();
        }

        /// <summary>
        /// Prompt for excel source
        /// Read excel source and store into data matrix
        /// Close source file
        /// </summary>
        public void ReadExcel(string path)
        {
            Debug.WriteLine($"Input file path: {path}");

            //Excel object reference to worksheet 1
            _excel = new Excel(path, 1);

            //Read source excel and store into data matrix
            _excel.ReadFile();

            //Close file
            _excel.Close();

            #region Debug
            //============================ DEBUG - PRINT OUT SOURCE DATA MATRIX================================
            Debug.WriteLine($"Num. of rows: {_excel.wsRowCount} | Num. of columns: {_excel.wsColumnCount}");
            for (int i = 0; i < _excel.wsRowCount; ++i)
            {
                for (int j = 0; j < _excel.wsColumnCount; ++j)
                {
                    Debug.Write($"{_excel.dataMatrix[i, j]} ");
                }
                Debug.WriteLine("");
            }
            #endregion
        }

        /// <summary>
        /// Build USPS AddressValidateRequest XML from Excel Source data matrix
        /// Deserialize each USPS Response
        /// Parse response to format and store Address and Carrier Route
        /// </summary>
        public void USPSAddressValidateRequest()
        {
            XmlSerializer x = new XmlSerializer(typeof(AddressValidateResponse));
            AddressValidateResponseAddress[] addressResponse;

            //Clear lists when new file loaded
            _parsedResponse.AddressPlusCarrier.Clear();
            _parsedResponse.CarrierPlusTally.Clear();

            #region Builder pattern for building each XML request

            for (int k = 0; k < _excel.wsRowCount; ++k)
            {
                _director.Construct(_builder, _excel, k);
                _bproduct = _builder.Retrieve();
                string xData = _webTool.AddressValidateReq(_bproduct);

                AddressValidateResponse myResponse = (AddressValidateResponse)x.Deserialize(new StringReader(xData));
                addressResponse = myResponse.Items;

                #region Debug
                //===================== DEBUG - PRINT OUT RESPONSE PARSED =================================
                Debug.WriteLine($"Address: {addressResponse[0].Address2}");
                Debug.WriteLine($"City: {addressResponse[0].City} | State: {addressResponse[0].State} | Zip: {addressResponse[0].Zip5}");
                Debug.WriteLine($"Carrier: {addressResponse[0].CarrierRoute}");
                #endregion

                //Store the response address and carrier route (Prepare for report 1)
                _parsedResponse.AddressPlusCarrier.Add(new string[2] { $"{addressResponse[0].Address2} {addressResponse[0].City} {addressResponse[0].State} {addressResponse[0].Zip5}", $"{addressResponse[0].CarrierRoute}" });

                //Store the carrier route and tally any repeats (Prepare for report 2)
                try
                {
                    if (_parsedResponse.CarrierPlusTally.ContainsKey(addressResponse[0].CarrierRoute))
                    {
                        _parsedResponse.CarrierPlusTally[addressResponse[0].CarrierRoute]++;
                    }
                    else
                    {
                        _parsedResponse.CarrierPlusTally.Add(addressResponse[0].CarrierRoute, 1);
                    }
                }
                catch(ArgumentNullException)
                {
                    if (_parsedResponse.CarrierPlusTally.ContainsKey("N/A"))
                    {
                        _parsedResponse.CarrierPlusTally["N/A"]++;
                    }
                    else
                    {
                        _parsedResponse.CarrierPlusTally.Add("N/A", 1);
                    }
                }
            }

            #endregion

            #region Validate req v1
            /*
            //Call AddressValidateRequest however many times there are # of rows in Excel Source
            for (int p = 0; p<_excel.wsRowCount; ++p)
            {
                //string xData = _webTool.AddressValidateRequest("Address1", "Address2", "City", "State", "Zip5", "Zip4");
                string xData = _webTool.AddressValidateRequest(
                    "",
                    $"{_excel.dataMatrix[p, 0]} {_excel.dataMatrix[p, 1]} {_excel.dataMatrix[p, 2]}",
                    $"{ _excel.dataMatrix[p, 3]}",
                    $"{ _excel.dataMatrix[p, 4]}",
                    $"{ _excel.dataMatrix[p, 5]}",
                    "");

                AddressValidateResponse myResponse = (AddressValidateResponse)x.Deserialize(new StringReader(xData));
                addressResponse = myResponse.Items;

                //===================== DEBUG - PRINT OUT RESPONSE PARSED =================================
                Debug.WriteLine($"Loop iteration: {p}");
                Debug.WriteLine($"Address: {addressResponse[0].Address2}");
                Debug.WriteLine($"City: {addressResponse[0].City} | State: {addressResponse[0].State} | Zip: {addressResponse[0].Zip5}");
                Debug.WriteLine($"Carrier: {addressResponse[0].CarrierRoute}");

                //Store the response address and carrier route (Prepare for report 1)
                _parsedResponse.AddressPlusCarrier.Add(new string[2] { $"{addressResponse[0].Address2} {addressResponse[0].City} {addressResponse[0].State} {addressResponse[0].Zip5}", $"{addressResponse[0].CarrierRoute}" });

                //Store the carrier route and tally any repeats (Prepare for report 2)
                if (_parsedResponse.CarrierPlusTally.ContainsKey(addressResponse[0].CarrierRoute))
                {
                    _parsedResponse.CarrierPlusTally[addressResponse[0].CarrierRoute]++;
                }
                else
                {
                    _parsedResponse.CarrierPlusTally.Add(addressResponse[0].CarrierRoute, 1);
                }
            }
            */
            #endregion

            Debug.WriteLine("AddressValidationComplete");
        }

        /// <summary>
        /// Calls factory method to create report one
        /// </summary>
        public void ExpReportOne()
        {
            Creator reportFactory = new ExporterOne();
            Product report = reportFactory.Create(_parsedResponse.AddressPlusCarrier, _excel);
            report.Save();
        }

        /// <summary>
        /// Calls factory method to create report two
        /// </summary>
        public void ExpReportTwo()
        {
            Creator reportFactory = new ExporterTwo();
            Product report = reportFactory.Create(_parsedResponse.CarrierPlusTally, _excel);
            report.Save();
        }
    }
}
