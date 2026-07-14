using System.Web.Script.Serialization;

namespace Ribbon.Vsto
{
    public static class JsonCodec
    {
        public static string Serialize(object value)
        {
            return new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.Serialize(value);
        }

        public static T Deserialize<T>(string json)
        {
            return new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.Deserialize<T>(json);
        }
    }
}
