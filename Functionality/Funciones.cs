using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ServicioConectorFE.Core;
using System.Xml.Serialization;
using System.Xml;

namespace ServicioConectorFE.Functionality
{
    public class Funciones
    {
        public string SerializeToXml<T>(T obj)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(T));

            StringWriter stringWriter = new StringWriter();

            XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces();
            namespaces.Add(string.Empty, string.Empty);

            XmlWriterSettings settings = new XmlWriterSettings
            {
                OmitXmlDeclaration = true,
                Indent = false
            };

            using (XmlWriter xmlWriter = XmlWriter.Create(stringWriter, settings))
            {
                serializer.Serialize(xmlWriter, obj, namespaces);
            }

            string xmlString = stringWriter.ToString();

            return xmlString;
        }

        public class XmlElementWrapper
        {
            public XmlElementWrapper(string name, object value)
            {
                Name = name;
                Value = value;
            }

            public string Name { get; set; }
            public object Value { get; set; }
        }

        #region DescargarArchivoURL
        public void downloadFileToSpecificPath(string strURLFile, string strPathToSave)
        {
            try
            {
                if (String.IsNullOrEmpty(strURLFile))
                {
                    throw new ArgumentNullException("La dirección URL del documento es nula o se encuentra en blanco.");
                }

                if (String.IsNullOrEmpty(strPathToSave))
                {
                    throw new ArgumentNullException("La ruta para almacenar el documento es nula o se encuentra en blanco.");
                }

                using (System.Net.WebClient client = new System.Net.WebClient())
                {
                    client.DownloadFile(strURLFile, strPathToSave);
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion
    }
}
