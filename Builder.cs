using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RouteMiner
{
    public abstract class Builder
    {
        public abstract void StreetNum(string s);
        public abstract void StreetName(string s);
        public abstract void AptNum(string s);
        public abstract void City(string s);
        public abstract void State(string s);
        public abstract void Zip(string s);
        public abstract void BuildRequestString();
        public abstract BuilderProduct Retrieve();
    }

    public class SBuilder : Builder
    {
        private BuilderProduct _product = new BuilderProduct();

        public override void StreetNum(string s) { _product.StreetNum = s; }
        public override void StreetName(string s) { _product.StreetName = s; }
        public override void AptNum(string s) { _product.AptNum = s; }
        public override void City(string s) { _product.City = s; }
        public override void State(string s) { _product.State = s; }
        public override void Zip(string s) { _product.Zip = s; }

        public override void BuildRequestString()
        {
            _product.ReqString = 
                $"{_product.StreetNum} " +
                $"{_product.StreetName} " +
                $"{_product.AptNum} " +
                $"{_product.City} " +
                $"{_product.State} " +
                $"{_product.Zip} "
                ;
        }

        public override BuilderProduct Retrieve()
        {
            return _product;
        }
    }

    /// <summary>
    /// Each product object would be 1 request ready for USPS Address Validate API
    /// </summary>
    public class BuilderProduct
    {
        private string _reqStr;
        private string _streetNum;
        private string _streetName;
        private string _aptNum;
        private string _city;
        private string _state;
        private string _zip;

        public string StreetNum { get { return _streetNum; } set { _streetNum = value; } }
        public string StreetName { get { return _streetName; } set { _streetName = value; } }
        public string AptNum { get { return _aptNum; } set { _aptNum = value; } }
        public string City { get { return _city; } set { _city = value; } }
        public string State { get { return _state; } set { _state = value; } }
        public string Zip { get { return _zip; } set { _zip = value; } }
        public string ReqString { get { return _reqStr; } set { _reqStr = value; } }
    }

    public class Director
    {
        public void Construct(Builder builder, Excel _excel, int i)
        {
            builder.StreetNum(_excel.dataMatrix[i, 0]);
            builder.StreetName(_excel.dataMatrix[i, 1]);
            builder.AptNum(_excel.dataMatrix[i, 2]);
            builder.City(_excel.dataMatrix[i, 3]);
            builder.State(_excel.dataMatrix[i, 4]);
            builder.Zip(_excel.dataMatrix[i, 5]);
        }
    }
}
