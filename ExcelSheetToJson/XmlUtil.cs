using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace ExcelSheetToJson
{
    public enum XmlUtilErrorCode
    {
        Found,
        NotFound
    }

    public class XmlUtil
    {
        private readonly Dictionary<string, XElement> dictIdXElements;

        public XmlUtil(Stream stream)
            : this()
        {
            Document = XDocument.Load(stream);
        }

        public XmlUtil(string xmlText)
            : this()
        {
            Document = XDocument.Parse(xmlText);
        }

        public XmlUtil(XmlDocument xmlDocument) : this()
        {
            Document = XDocument.Parse(xmlDocument.OuterXml);
        }

        private XmlUtil()
        {
            dictIdXElements = new Dictionary<string, XElement>();
            SplitCharacter = '.';
            Functions = new List<XmlUtilFunc>();
        }

        internal string GetAttributeValue(XElement xElement, string attribute)
        {
            var foundAttribute = xElement.Attributes().FirstOrDefault(x => x.Name.LocalName == attribute);
            return foundAttribute?.Value;
        }

        public XDocument Document { get; }
        public IList<XmlUtilFunc> Functions { get; set; }
        public char SplitCharacter { get; set; }

        public XmlDocument XmlDocument
        {
            get
            {
                var xmlDocument = new XmlDocument();
                using (var xmlReader = Document.CreateReader())
                {
                    xmlDocument.LoadXml(Document.ToString());
                }
                return xmlDocument;
            }
        }

        public XmlUtilResult Get(string path)
        {
            foreach (var func in Functions)
            {
                if (func.CanExecute(path))
                    return func.Execute(path);
            }

            return GetElement(path);
        }

        public IEnumerable<XElement> GetElements(string path)
        {
            var splitCharacterIndex = path.LastIndexOf(SplitCharacter);
            if (splitCharacterIndex > 0)
            {
                var parentPath = path.Substring(0, splitCharacterIndex);
                var parentXElement = Get(parentPath).Element;
                if (parentXElement == null) return new List<XElement>();
                var element = path.Substring(splitCharacterIndex + 1);
                return
                    parentXElement.Descendants()
                        .Where(x => x.Name.LocalName.Equals(element, StringComparison.OrdinalIgnoreCase))
                        .ToList();
            }
            return Document.Root.Descendants()
                .Where(x => x.Name.LocalName.Equals(path, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        public XElement GetRelativeElement(XElement xElement, string path)
        {
            var paths = path.Split(SplitCharacter);
            var result = xElement;
            foreach (var item in paths)
            {
                result =
                    result?.Descendants()
                        .FirstOrDefault(x => x.Name.LocalName.Equals(item, StringComparison.OrdinalIgnoreCase));

                if (result == null)
                    break;
            }

            return result;
        }

        public void SetRelativeElementValue(XElement xElement, string path, string value)
        {
            var foundXElement = GetRelativeElement(xElement, path);
            foundXElement?.SetValue(value);
        }

        public void SetValue(string path, string value)
        {
            var element = Get(path).Element;
            element?.SetValue(value);
        }

        internal string GetAttributeValue(string path, string attribute)
        {
            var result = GetElement(path);
            if (result.ErrorCode != XmlUtilErrorCode.Found) return null;
            var attributeElement =
                result.Element.Attributes().ToList().FirstOrDefault(x => x.ToString().StartsWith(attribute));
            return attributeElement?.Value;
        }

        internal void Write(string fileFullPath)
        {
            File.WriteAllText(fileFullPath, Document.ToString());
        }

        private XmlUtilResult GetElement(string path)
        {
            var elements = path.Split(SplitCharacter);
            var element = Document.Root;

            foreach (var item in elements)
            {
                if (element == null) continue;
                var regex = new Regex(@"(\w+)?\[(\d+)\]");
                if (regex.IsMatch(item))
                {
                    var match = regex.Match(item);
                    element =
                        element.Elements()
                            .Where(
                                x =>
                                    x.Name.LocalName.Equals(match.Groups[1].Value,
                                        StringComparison.OrdinalIgnoreCase))
                            .ToList()[Convert.ToInt32(match.Groups[2].Value)];
                }
                else
                {
                    element = element.Elements()
                        .FirstOrDefault(x => x.Name.LocalName.Equals(item, StringComparison.OrdinalIgnoreCase));
                }

                var refAttribute = element?.Attributes()
                    .FirstOrDefault(x => x.Name.LocalName.Equals("Ref", StringComparison.OrdinalIgnoreCase));
                if (refAttribute != null)
                    element = GetElementWithId(refAttribute.Value);
            }
            return element != null
                ? new XmlUtilResult {ErrorCode = XmlUtilErrorCode.Found, Value = element.Value, Element = element}
                : new XmlUtilResult {ErrorCode = XmlUtilErrorCode.NotFound};
        }

        private XElement GetElementWithId(string id)
        {
            var key = id;
            if (dictIdXElements.ContainsKey(key)) return dictIdXElements[key];
            var elements = Document.Descendants().Where(
                x => x.Attributes().Any(y => y.Name.LocalName.Equals("Id", StringComparison.OrdinalIgnoreCase)));

            foreach (var item in elements)
            {
                var refAttribute = item.Attributes()
                    .FirstOrDefault(x => x.Name.LocalName.Equals("Id", StringComparison.OrdinalIgnoreCase));
                if (refAttribute != null && refAttribute.Value.Equals(id, StringComparison.OrdinalIgnoreCase))
                {
                    if (!dictIdXElements.ContainsKey(key))
                        dictIdXElements.Add(key, item);
                    return item;
                }
            }
            return null;
        }

        public IList<XElement> GetRelativeElements(XElement xElement, string path)
        {
            var result = new List<XElement>();
            var paths = path.Split(SplitCharacter);
            var nextPath = paths[0];
            var remainingPath = string.Join(".", paths,1, paths.Length-1);
            var childs = xElement.Descendants()
                .Where(x => x.Name.LocalName.Equals(nextPath, StringComparison.OrdinalIgnoreCase)).ToList();
            if (string.IsNullOrWhiteSpace(remainingPath))
            {
                result.AddRange(childs);
            }
            else
            {
                foreach (var element in childs)
                {
                    result.AddRange(GetRelativeElements(element, remainingPath));
                }
            }
            return result;
        }
    }

    public class XmlUtilFunc
    {
        public Func<string, bool> CanExecute { get; set; }
        public Func<string, XmlUtilResult> Execute { get; set; }
    }

    public class XmlUtilResult
    {
        public bool? BooleanNullableValue
        {
            get
            {
                bool result;
                if (bool.TryParse(Value, out result))
                    return result;
                return null;
            }
        }

        public DateTime? DateTimeNullableValue
        {
            get
            {
                DateTime result;
                if (DateTime.TryParse(Value, out result))
                    return result;
                return null;
            }
        }

        public decimal? DecimalNullableValue
        {
            get
            {
                decimal result;
                if (decimal.TryParse(Value, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
                    return result;
                return null;
            }
        }

        public double? DoubleNullableValue
        {
            get
            {
                double result;
                if (double.TryParse(Value, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
                    return result;
                return null;
            }
        }

        public XElement Element { get; set; }
        public XmlUtilErrorCode ErrorCode { get; set; }
        public object ResultObject { get; set; }
        public string Value { get; set; }
    }
}