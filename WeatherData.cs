using System.Net;
using System.Xml;

namespace Voice_Bot
{
    class WeatherData
    {
        //private only properties commonly use for property that we don't expose them to the public e.g Data Repository, Domain Services

        //Property accessors (getter and setter)
        public string City { get; set; }
        public float Temperature { get; set; }
        public string Condition { get; set; }

        public WeatherData(string city)
        {
            City = city;
        }

        //Update weather
        public void CheckWeather()
        {
            WeatherAPI DataAPI = new WeatherAPI(City);
            Temperature = DataAPI.GetTemp();
            Condition = DataAPI.GetCondition();
        }
    }

    class WeatherAPI
    {
        //You need to create your own APIKEY in Openweathermap if you want to use its API
        private const string APIKEY = "YOUR APIKEY HERE";
        private string CurrentURL;
        private XmlDocument _xmlDocument;

        public WeatherAPI(string city)
        {
            SetCurrentURL(city);
            _xmlDocument = GetXML(CurrentURL);
        }

        public float GetTemp()
        {
            //Get value in temperature
            XmlNode temp_node = _xmlDocument.SelectSingleNode("//temperature");
            XmlAttribute temp_value = temp_node.Attributes["value"];
            string temp_string = temp_value.Value;

            return float.Parse(temp_string);
        }

        public string GetCondition()
        {
            //Get value in weather
            XmlNode condition_node = _xmlDocument.SelectSingleNode("//weather");
            XmlAttribute condition_value = condition_node.Attributes["value"];
            string condition_string = condition_value.Value;

            return condition_string;
        }

        private void SetCurrentURL(string location)
        {
            /**
             * If you have an APIKEY, you can open this link below to check current weather data.
             * Remember to assign location (eg. Auckland) and APIKEY.
             */
            CurrentURL = "http://api.openweathermap.org/data/2.5/weather?q=" 
                + location + "&mode=xml&units=metric&APPID=" + APIKEY;
        }

        private XmlDocument GetXML(string CurrentURL)
        {
            using (WebClient client = new WebClient())
            {
                string xmlContent = client.DownloadString(CurrentURL);
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(xmlContent);
                return xmlDocument;
            }
        }
    }
}
