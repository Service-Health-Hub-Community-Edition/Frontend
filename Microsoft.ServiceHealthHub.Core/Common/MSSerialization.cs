/* using Microsoft.Graph;
using Microsoft.IdentityModel.Tokens;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Microsoft.ServiceHealthHub.Core
{
    /// <summary>
    /// Persisted attribute class. Fields in classes inheriting from
    /// ZTAutoSerializingObject marked with this attribute will be
    /// serialized and deserialized as needed.
    /// </summary>
    /// <owner alias="aldras" />
    [Guid("5FB58511-24F2-463E-8692-8428D9559993")]
    public class PersistedAttribute : Attribute
    {

    }

    /// <summary>
    /// Secure Store Persisted attribute class. Fields in classes inheriting from
    /// ZTAutoSerializingObject marked with this attribute will be
    /// stored and retrieved to / from central Key Vault instance as needed.
    /// </summary>
    /// <owner alias="aldras" />
    [Guid("271EDC5A-99FF-4885-8D5F-E22DDD26D888")]
    public class SecureStorePersistedAttribute : Attribute
    {

    }

    /// <summary>
    /// Provides automatic serialization of the marked fields. Used as
    /// a base class for the ZTBasePersisted class which is base class for
    /// all configuration classes
    /// </summary>
    /// <owner alias="aldras" />
    [Guid("3073EB37-3483-4438-A597-422919D78D0C")]
    public class CHAutoSerializingObject
    {
        #region Constants
        internal const BindingFlags bindingFlags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly;
        #endregion Constants

        #region Fields
        /// <summary>
        /// Internal store for all fields which cannot be deserialized.
        /// During the deserialization of the object, the xml structure passed
        /// to the deserializer may contain fields which do not exist in the 
        /// class anymore. Those fields and values will be stored in this hashtable.
        /// During the serialization of the object, the items from this hashtable
        /// will be serialized to the old location so that the structure of the 
        /// object record in the configuration database is not altered and
        /// no value is lost.
        /// </summary>
        /// <seealso cref="UpgradedPersistedProperties"/>
        /// <owner alias="aldras" />
        private Hashtable _upgradedPersistedProperties = new Hashtable();
        #endregion Fields

        #region Properties
        /// <summary>
        /// Public property for accessing for the fields which could not be deserialized.
        /// </summary>
        /// <seealso cref="_upgradedPersistedProperties"/>
        /// <owner alias="aldras" />
        public Hashtable UpgradedPersistedProperties
        {
            get { return _upgradedPersistedProperties; }
        }
        #endregion Properties

        #region Constructors
        /// <summary>
        /// Base ZTAutoSerializingObject constructor. For internal use only.
        /// </summary>
        /// <owner alias="aldras" />
        protected CHAutoSerializingObject()
        {
        }
        #endregion Constructors

        #region Helper methods
        /// <summary>
        /// Enumerates all the fields in the given class type declaration and all base classes
        /// in the inheritance structure.
        /// </summary>
        /// <param name="type">
        /// Type declaration which needs to be scanned.
        /// </param>
        /// <seealso cref="DeserializeBasicObject(string)"/>
        /// <seealso cref="SerializeBasicObject"/>;
        /// <owner alias="aldras" />
        internal Dictionary<string, FieldInfo> EnumFields(Type type)
        {
            Dictionary<string, FieldInfo> fields = new Dictionary<string, FieldInfo>();
            if (type.BaseType != null)
            {
                Dictionary<string, FieldInfo> pFields = EnumFields(type.BaseType);
                foreach (KeyValuePair<string, FieldInfo> f in pFields)
                    fields[f.Key] = f.Value;
            }

            FieldInfo[] fldCol = type.GetFields(bindingFlags);
            foreach (FieldInfo fld in fldCol)
            {
                Attribute attr = Attribute.GetCustomAttribute(fld, typeof(PersistedAttribute), true);
                if (attr != null)
                {
                    fields[fld.Name] = fld;
                }
            }

            return fields;
        }
        #endregion Helper methods

        #region Deserialization
        /// <summary>
        /// Deserializes the object based on the data provided in the XML structure.
        /// Intended for internal use only.
        /// </summary>
        /// <param name="xmlDefinition">
        /// Object data in the XML format
        /// </param>
        /// <seealso cref="ParseObject(string, string)"/>
        /// <seealso cref="ParseObject(string, XmlNode)"/>
        /// <owner alias="aldras" />
        public void DeserializeBasicObject(string xmlDefinition)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xmlDefinition);
            string assemblyQName = xmlDoc.FirstChild.Attributes["Type"].Value;

            // ZTTraceLogging.LogEvent(ZTLoggingLevel.Information, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00100", string.Format("Deserialization started for the object type: {0}.", this.GetType().AssemblyQualifiedName), "");

            if (assemblyQName != this.GetType().AssemblyQualifiedName)
            {
                ZTTraceLogging.LogEvent(ZTLoggingLevel.Error, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00101", string.Format("Deserialization failed with the following error: Type mismatch. Object type: {0}, database object type: {1}", this.GetType().AssemblyQualifiedName, assemblyQName), "");
                throw new Exception(string.Format("Type mismatch. Object type: {0}, database object type: {1}", this.GetType().AssemblyQualifiedName, assemblyQName));
            }

            Dictionary<string, FieldInfo> fields = EnumFields(this.GetType());

            object o;
            string Name;
            string Type;
            string Value;
            string xmlValue;

            // TODO: hardcode object and fld elements here
            foreach (XmlNode xmlNode in xmlDoc.FirstChild.ChildNodes)
            {
                Name = "";
                Type = "";
                Value = "";
                xmlValue = "";

                try
                {
                    Name = xmlNode.Attributes["Name"].Value;
                    Type = xmlNode.Attributes["Type"].Value;
                    Value = xmlNode.InnerText;
                    xmlValue = xmlNode.InnerXml;
                }
                catch
                {
                    if (string.IsNullOrEmpty(Name))
                    {
                        Name = string.Format("Unknown-{0}", Guid.NewGuid());
                        ZTTraceLogging.LogEvent(ZTLoggingLevel.Warning, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00104", string.Format("Deserialization: Property name is not provided. Following property name will be used: \"{0}\".", Name), "");
                    }

                    if (string.IsNullOrEmpty(Type))
                    {
                        Type = "System.String";
                        ZTTraceLogging.LogEvent(ZTLoggingLevel.Warning, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00105", string.Format("Deserialization: Property type is not provided. Following property type will be used: \"{0}\".", Type), "");
                    }

                    try
                    {
                        Value = xmlNode.InnerText;
                    }
                    catch
                    {
                        Value = "";
                        ZTTraceLogging.LogEvent(ZTLoggingLevel.Warning, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00106", "Deserialization: Property value is not provided. Using an empty string.", "");
                    }
                }

                if (xmlNode.SelectNodes("cfld").Count > 0)
                    o = ParseObject(Type, xmlNode);
                else
                    o = ParseObject(Type, Value, xmlValue);

                if (fields.ContainsKey(Name) && null != fields[Name])
                {
                    fields[Name].SetValue(this, o);
                }
                else
                {
                    _upgradedPersistedProperties[Name] = o;
                }
            }

            // ZTTraceLogging.LogEvent(ZTLoggingLevel.Information, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00111", string.Format("Deserialization completed for the object type: {0}. {1} properties found.", this.GetType().AssemblyQualifiedName, xmlDoc.FirstChild.ChildNodes.Count.ToString()), "");
        }

        private ZTBasePersistedObject InitializeBasePersistedObject(string id, string classType = "")
        {
            ZTBasePersistedObject obj = null;
            if (Guid.TryParse(id, out Guid objectId))
            {
                try
                {
                    obj = ZTObjectCache.Instance.GetPersistedObject(objectId);
                }
                catch (Exception ex)
                {
                    ZTTraceLogging.LogEvent(ZTLoggingLevel.Error, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO001A1", $"Cannot initialize the object with the ID {id}. Exception: {ex}", "");
                }
            }

            return obj;
        }

        /// <summary>
        /// Return an object based on the provided data type and the value.
        /// Intended for internal use only.
        /// </summary>
        /// <param name="Type">
        /// Object type as a string. Following types are supported:
        /// System.Guid, System.String, System.Bool,
        /// System.Int16, System.Int32, System.Int64,
        /// System.UInt16, System.UInt32, System.UInt64,
        /// System.Byte, System.DateTime, System.Uri,
        /// Microsoft.ZeroTouch.Administration.ZTLoggingLevel,
        /// Microsoft.ZeroTouch.Administration.ZTServerRole
        /// </param>
        /// <param name="Value">
        /// Object value as string. Based on the provided object type,
        /// the value will be casted and return as requested object
        /// </param>
        /// <seealso cref="ParseObject(string, XmlNode)"/>
        /// <seealso cref="DeserializeBasicObject(string)"/>
        /// <owner alias="aldras" />
        private object ParseObject(string Type, string Value, string xmlValue)
        {
            object o = null;

            try
            {
                switch (Type)
                {
                    case "System.Guid":
                        o = new Guid(Value);
                        break;
                    case "System.String":
                        o = Value;
                        break;
                    case "System.Boolean":
                        o = Convert.ToBoolean(Value);
                        break;
                    case "System.Int16":
                        o = Convert.ToInt16(Value);
                        break;
                    case "System.Int32":
                        o = Convert.ToInt32(Value);
                        break;
                    case "System.Int64":
                        o = Convert.ToInt64(Value);
                        break;
                    case "System.UInt16":
                        o = Convert.ToUInt16(Value);
                        break;
                    case "System.UInt32":
                        o = Convert.ToUInt32(Value);
                        break;
                    case "System.UInt64":
                        o = Convert.ToUInt64(Value);
                        break;
                    case "System.Double":
                        o = Convert.ToDouble(Value, CultureInfo.InvariantCulture);
                        break;
                    case "System.Byte":
                        o = Convert.ToByte(Value);
                        break;
                    case "System.DateTime":
                        o = Convert.ToDateTime(Value);
                        break;
                    case "System.Version":
                        o = Version.Parse(Value);
                        break;
                    case "System.Uri":
                        o = new Uri(Value);
                        break;
                    case "System.Text.Encoding":
                        o = Encoding.GetEncoding(Value);
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTLoggingLevel":
                        o = (ZTLoggingLevel)Enum.Parse(typeof(ZTLoggingLevel), Value);
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTServerRole":
                        o = (ZTServerRole)Enum.Parse(typeof(ZTServerRole), Value);
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTJobLockType":
                        o = (ZTJobLockType)Enum.Parse(typeof(ZTJobLockType), Value);
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTWebApplicationType":
                        o = (ZTWebApplicationType)Enum.Parse(typeof(ZTWebApplicationType), Value);
                        break;
                    case "System.Collections.Hashtable":
                        // just making sure that there is a collection initialized, even if it is empty.
                        o = new Hashtable();
                        break;
                    case "System.Byte[]":
                        // just making sure that there is a collection initialized, even if it is empty.
                        o = new byte[0];
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTEncryptedString":
                        // if the code path ends here, it means the collection is empty within the object xml.
                        // in this case, we are returning null
                        o = null;
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTServer":
                        Guid serverGuid;
                        if (Guid.TryParse(Value, out serverGuid))
                        {
                            try
                            {
                                o = new ZTServer(serverGuid);
                            }
                            catch
                            {
                                // log error
                                o = null;
                            }
                            // ZTTraceLogging.LogEvent(ZTLoggingLevel.Information, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00107", string.Format("Server object created successfully. Type: \"{0}\", id: \"{1}\", name:\"{2}\".", Type, Value, ((ZTServer)o).Name), "");
                        }
                        else
                        {
                            // ZTTraceLogging.LogEvent(ZTLoggingLevel.Warning, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00108", string.Format("Server guid invalid or not present. Type: \"{0}\", Value: \"{1}\". Returning empty object.", Type, Value), "");
                            o = null;
                        }
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTEndpoint":
                        Guid endpointGuid;
                        if (Guid.TryParse(Value, out endpointGuid))
                        {
                            try
                            {
                                ZTBasePersistedObject obj = ZTObjectCache.Instance.GetPersistedObject(endpointGuid);
                                if (obj is ZTEndpoint)
                                    o = obj;
                                else
                                    o = null;
                            }
                            catch
                            {
                                // log error
                                o = null;
                            }
                            // ZTTraceLogging.LogEvent(ZTLoggingLevel.Information, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00107", string.Format("Server object created successfully. Type: \"{0}\", id: \"{1}\", name:\"{2}\".", Type, Value, ((ZTServer)o).Name), "");
                        }
                        else
                        {
                            // ZTTraceLogging.LogEvent(ZTLoggingLevel.Warning, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00108", string.Format("Server guid invalid or not present. Type: \"{0}\", Value: \"{1}\". Returning empty object.", Type, Value), "");
                            o = null;
                        }
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTDatabase":
                        o = InitializeBasePersistedObject(Value, Type);
                        break;
                    case "Microsoft.ZeroTouch.Connectivity.Orchestrator.ZTOrchestratorDatabase":
                        o = InitializeBasePersistedObject(Value, Type);
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTNotificationEndpoint":
                        o = InitializeBasePersistedObject(Value, Type);
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTNotificationEndpointCollection":
                        // just making sure that there is a collection initialized, even if it is empty.
                        o = new ZTNotificationEndpointCollection();
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTPowerShellScript":
                        Guid psScriptGuid;
                        if (Guid.TryParse(Value, out psScriptGuid))
                        {
                            try
                            {
                                o = new ZTPowerShellScript(psScriptGuid);
                            }
                            catch
                            {
                                // log error
                                o = null;
                            }
                            // ZTTraceLogging.LogEvent(ZTLoggingLevel.Information, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00107", string.Format("Server object created successfully. Type: \"{0}\", id: \"{1}\", name:\"{2}\".", Type, Value, ((ZTServer)o).Name), "");
                        }
                        else
                        {
                            // ZTTraceLogging.LogEvent(ZTLoggingLevel.Warning, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00108", string.Format("Server guid invalid or not present. Type: \"{0}\", Value: \"{1}\". Returning empty object.", Type, Value), "");
                            o = null;
                        }
                        break;
                    case "Microsoft.ZeroTouch.ZTSchedule":
                        if (string.IsNullOrEmpty(Value))
                            o = null;
                        else
                            o = ZTSchedule.Parse(Value);
                        break;
                    case "System.Management.Automation.Runspaces.AuthenticationMechanism":
                        o = (AuthenticationMechanism)Enum.Parse(typeof(AuthenticationMechanism), Value);
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTBaseActionParameterDefinitionCollection":
                        o = new ZTBaseActionParameterDefinitionCollection(Value);
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTBaseActionParameterDefinition":
                        try
                        {
                            o = new ZTBaseActionParameterDefinition(Value);
                        }
                        catch
                        {
                            o = null;
                        }
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTTicketApprovedConditionCollection":
                        o = new ZTTicketApprovedConditionCollection(Value);
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTKillSwitch":
                        if (string.IsNullOrWhiteSpace(Value))
                            o = null;
                        else
                        {
                            ZTKillSwitch ks = new ZTKillSwitch();
                            ks.Deserialize(Value);
                            o = ks;
                        }
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTTicketApprovedCondition":
                        try
                        {
                            o = new ZTTicketApprovedCondition(Value);
                        }
                        catch
                        {
                            o = null;
                        }
                        break;
                    case "Microsoft.ZeroTouch.Administration.ZTHealthCheckResult":
                        o = new ZTHealthCheckResult();
                        ((ZTHealthCheckResult)o).Deserialize(Value);
                        break;
                    default:
                        o = null;
                        bool deserialized = false;
                        List<Type> fullTypeList = ZTUtility.GetTypeByName(Type);
                        if (fullTypeList.Count > 0)
                        {
                            Type type = fullTypeList[0];
                            object obj = Activator.CreateInstance(type, true);
                            if (obj is IZTSerializable)
                            {
                                ((IZTSerializable)obj).Deserialize(xmlValue);
                                o = obj;
                                deserialized = true;
                            }
                        }

                        if (!deserialized)
                        {
                            o = null;
                            try
                            {
                                o = InitializeBasePersistedObject(Value, Type);
                            }
                            catch
                            {
                                // just ignore the exception if the object is not found
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                ZTTraceLogging.LogEvent(ZTLoggingLevel.Warning, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00109", string.Format("Deserialization failed due to the invalid data provided. Type: \"{0}\", Value: \"{1}\". Exception: {2}. Returning empty object.", Type, Value, ex), "");
                o = null;
            }

            return o;
        }

        /// <summary>
        /// Return a collection of objects based on the provided data type and the value.
        /// Intended for internal use only.
        /// </summary>
        /// <param name="Type">
        /// Collection type as a string. Following collections are supported:
        /// System.Collections.Hashtable, System.Byte[]
        /// </param>
        /// <param name="Value">
        /// Collection data as XmlNode. Based on the provided collection type,
        /// the XmlNode is processed, collection created and returned as an output.
        /// </param>
        /// <seealso cref="ParseObject(string, string)"/>
        /// <seealso cref="DeserializeBasicObject(string)"/>
        /// <owner alias="aldras" />
        private object ParseObject(string Type, XmlNode Value)
        {
            object o;

            try
            {
                switch (Type)
                {
                    case "System.Collections.Hashtable":
                        {
                            Hashtable ht = new Hashtable();
                            XmlNodeList xmlNl = Value.SelectNodes("cfld");

                            object colPropObj;
                            string colPropName;
                            string colPropType;
                            string colPropValue;
                            string colPropXmlValue;

                            foreach (XmlNode xmlNode in xmlNl)
                            {
                                colPropName = (null != xmlNode.Attributes["Name"]) ? xmlNode.Attributes["Name"].Value : colPropName = string.Format("Unknown-{0}", Guid.NewGuid());
                                colPropType = (null != xmlNode.Attributes["Type"]) ? xmlNode.Attributes["Type"].Value : colPropType = "System.String";
                                colPropValue = xmlNode.InnerText;
                                colPropXmlValue = xmlNode.InnerXml;

                                colPropObj = ParseObject(colPropType, colPropValue, colPropXmlValue);

                                ht[colPropName] = colPropObj;
                            }

                            o = ht;
                            break;
                        }
                    case "System.Byte[]":
                        {
                            XmlNodeList xmlNl = Value.SelectNodes("cfld");

                            byte[] byteArray = new byte[xmlNl.Count];

                            object colPropObj;
                            string colPropName;
                            string colPropType;
                            string colPropValue;
                            string colPropXmlValue;
                            int c = 0;

                            foreach (XmlNode xmlNode in xmlNl)
                            {
                                colPropName = (null != xmlNode.Attributes["Name"]) ? xmlNode.Attributes["Name"].Value : colPropName = string.Format("Unknown-{0}", Guid.NewGuid());
                                colPropType = (null != xmlNode.Attributes["Type"]) ? xmlNode.Attributes["Type"].Value : colPropType = "System.String";
                                colPropValue = xmlNode.InnerText;
                                colPropXmlValue = xmlNode.InnerXml;

                                colPropObj = ParseObject(colPropType, colPropValue, colPropXmlValue);

                                if (colPropObj.GetType() == typeof(byte))
                                {
                                    byteArray[c] = (byte)colPropObj;
                                }

                                c++;
                            }

                            o = byteArray;
                            break;
                        }
                    case "Microsoft.ZeroTouch.Administration.ZTEncryptedString":
                        {
                            XmlNodeList xmlNl = Value.SelectNodes("cfld");

                            byte[] byteArray = new byte[xmlNl.Count];

                            object colPropObj;
                            string colPropName;
                            string colPropType;
                            string colPropValue;
                            string colPropXmlValue;
                            int c = 0;

                            foreach (XmlNode xmlNode in xmlNl)
                            {
                                colPropName = (null != xmlNode.Attributes["Name"]) ? xmlNode.Attributes["Name"].Value : colPropName = string.Format("Unknown-{0}", Guid.NewGuid());
                                colPropType = (null != xmlNode.Attributes["Type"]) ? xmlNode.Attributes["Type"].Value : colPropType = "System.String";
                                colPropValue = xmlNode.InnerText;
                                colPropXmlValue = xmlNode.InnerXml;

                                colPropObj = ParseObject(colPropType, colPropValue, colPropXmlValue);

                                if (colPropObj.GetType() == typeof(byte))
                                {
                                    byteArray[c] = (byte)colPropObj;
                                }

                                c++;
                            }

                            o = new ZTEncryptedString(byteArray);
                            break;
                        }
                    case "Microsoft.ZeroTouch.Administration.ZTNotificationEndpointCollection":
                        {
                            ZTNotificationEndpointCollection col = new ZTNotificationEndpointCollection();
                            XmlNodeList xmlNl = Value.SelectNodes("cfld");

                            object colPropObj;
                            string colPropName;
                            string colPropType;
                            string colPropValue;

                            foreach (XmlNode xmlNode in xmlNl)
                            {
                                colPropName = (null != xmlNode.Attributes["Name"]) ? xmlNode.Attributes["Name"].Value : colPropName = string.Format("Unknown-{0}", Guid.NewGuid());
                                colPropType = (null != xmlNode.Attributes["Type"]) ? xmlNode.Attributes["Type"].Value : colPropType = "Microsoft.ZeroTouch.Administration.ZTNotificationEndpoint";
                                colPropValue = xmlNode.InnerText;

                                Guid g;
                                if (Guid.TryParse(colPropValue, out g))
                                    colPropObj = InitializeBasePersistedObject(colPropValue, colPropType);
                                else
                                    colPropObj = null;

                                if (colPropObj != null && colPropObj is ZTNotificationEndpoint)
                                    col[colPropName] = (ZTNotificationEndpoint)colPropObj;
                            }

                            o = col;
                            break;
                        }
                    case "System.String":
                        o = Value.InnerXml;
                        break;
                    default:
                        o = null;
                        break;
                }
            }
            catch (Exception ex)
            {
                // log conversion failed
                ZTTraceLogging.LogEvent(ZTLoggingLevel.Warning, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00114", string.Format("Deserialization failed due to the invalid data provided. Type: \"{0}\", Value: \"{1}\". Exception: {2}. Returning empty object.", Type, Value, ex), "");
                o = null;
            }

            return o;
        }
        #endregion Deserialization

        #region Serialization
        /// <summary>
        /// Serializes the collection and returns the XmlElement based on the collection data and type.
        /// Intended for internal use only.
        /// </summary>
        /// <param name="node">Root node under which the XmlElement should be created</param>
        /// <param name="Name">Name of the collection being serialized</param>
        /// <param name="Type">
        /// Collection type as a string. Following collections are supported:
        /// System.Collections.Hashtable, System.Byte[]
        /// </param>
        /// <param name="o">Collection object to serialize</param>
        /// <seealso cref="SerializeField(XmlElement, string, string, string, object)"/>
        /// <owner alias="aldras" />
        XmlElement SerializeCollection(XmlElement node, string Name, string Type, object o)
        {
            XmlElement xmlElement = node.OwnerDocument.CreateElement("fld");
            if (Name != string.Empty)
                xmlElement.SetAttribute("Name", Name);
            xmlElement.SetAttribute("Type", Type);

            if (null != o)
            {
                switch (Type)
                {
                    case "System.Collections.Hashtable":
                        {
                            Hashtable ht = (Hashtable)o;
                            foreach (DictionaryEntry item in ht)
                            {
                                XmlElement xmlFld = SerializeField(node, "cfld", item.Key.ToString(), item.Value.GetType().FullName, item.Value);
                                xmlElement.AppendChild(xmlFld);
                            }
                            break;
                        }
                    case "System.Byte[]":
                        {
                            byte[] byteArray = (byte[])o;
                            foreach (byte item in byteArray)
                            {
                                XmlElement xmlFld = SerializeField(node, "cfld", string.Empty, item.GetType().FullName, item);
                                xmlElement.AppendChild(xmlFld);
                            }
                            break;
                        }
                    case "Microsoft.ZeroTouch.Administration.ZTEncryptedString":
                        {
                            byte[] byteArray = ((ZTEncryptedString)o).EncryptedString;
                            foreach (byte item in byteArray)
                            {
                                XmlElement xmlFld = SerializeField(node, "cfld", string.Empty, item.GetType().FullName, item);
                                xmlElement.AppendChild(xmlFld);
                            }
                            break;
                        }
                    case "Microsoft.ZeroTouch.Administration.ZTNotificationEndpointCollection":
                        {
                            ZTNotificationEndpointCollection col = (ZTNotificationEndpointCollection)o;
                            foreach (KeyValuePair<string, ZTNotificationEndpoint> item in col)
                            {
                                XmlElement xmlFld = SerializeField(node, "cfld", item.Key.ToString(), item.Value.GetType().FullName, item.Value.Id);
                                xmlElement.AppendChild(xmlFld);
                            }
                            break;
                        }
                    default:
                        xmlElement.InnerText = o.ToString();
                        break;
                }
            }
            return xmlElement;
        }

        private string SerializePersistedObject(ZTBasePersistedObject obj)
        {
            if (obj != null)
                return obj.Id.ToString();
            else
                return "";
        }

        /// <summary>
        /// Serializes the object and returns the XmlElement based on the object data and type.
        /// Intended for internal use only.
        /// </summary>
        /// <param name="node">Root node under which the XmlElement should be created</param>
        /// <param name="elementName">Name of the XML element. Valid values are: fld for the field, sfld for the collection members</param>
        /// <param name="Name">Name of the field which is being serialized</param>
        /// <param name="Type">Type of field. Valid values are:
        /// System.Guid, System.String, System.Bool,
        /// System.Int16, System.Int32, System.Int64,
        /// System.UInt16, System.UInt32, System.UInt64,
        /// System.Byte, System.DateTime, System.Uri,
        /// Microsoft.ZeroTouch.Administration.ZTLoggingLevel,
        /// Microsoft.ZeroTouch.Administration.ZTServerRole 
        /// </param>
        /// <param name="o">Object to serialize</param>
        /// <seealso cref="SerializeBasicObject"/>
        /// <seealso cref="SerializeCollection(XmlElement, string, string, object)"/>
        /// <owner alias="aldras" />
        private XmlElement SerializeField(XmlElement node, string elementName, string Name, string Type, object o)
        {
            XmlElement xmlElement = node.OwnerDocument.CreateElement(elementName);
            if (Name != string.Empty)
                xmlElement.SetAttribute("Name", Name);

            xmlElement.SetAttribute("Type", Type);

            switch (Type)
            {
                case "System.Guid":
                    xmlElement.InnerText = o.ToString();
                    break;
                case "System.String":
                    xmlElement.InnerText = (string)o;
                    break;
                case "System.Boolean":
                    xmlElement.InnerText = o.ToString();
                    break;
                case "System.Int16":
                    xmlElement.InnerText = o.ToString();
                    break;
                case "System.Int32":
                    xmlElement.InnerText = o.ToString();
                    break;
                case "System.Int64":
                    xmlElement.InnerText = o.ToString();
                    break;
                case "System.UInt16":
                    xmlElement.InnerText = o.ToString();
                    break;
                case "System.UInt32":
                    xmlElement.InnerText = o.ToString();
                    break;
                case "System.UInt64":
                    xmlElement.InnerText = o.ToString();
                    break;
                case "System.Double":
                    xmlElement.InnerText = ((double)o).ToString(CultureInfo.InvariantCulture);
                    break;
                case "System.Byte":
                    xmlElement.InnerText = o.ToString();
                    break;
                case "System.DateTime":
                    xmlElement.InnerText = ((DateTime)o).ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
                    break;
                case "System.Version":
                    xmlElement.InnerText = o.ToString();
                    break;
                case "System.Uri":
                    xmlElement.InnerText = ((Uri)o).ToString();
                    break;
                case "System.Text.Encoding":
                    xmlElement.InnerText = ((Encoding)o).WebName;
                    break;
                case "Microsoft.ZeroTouch.Administration.ZTLoggingLevel":
                    xmlElement.InnerText = ((ZTLoggingLevel)o).ToString();
                    break;
                case "Microsoft.ZeroTouch.Administration.ZTServerRole":
                    xmlElement.InnerText = ((ZTServerRole)o).ToString();
                    break;
                case "Microsoft.ZeroTouch.Administration.ZTJobLockType":
                    xmlElement.InnerText = ((ZTJobLockType)o).ToString();
                    break;
                case "Microsoft.ZeroTouch.Administration.ZTWebApplicationType":
                    xmlElement.InnerText = ((ZTWebApplicationType)o).ToString();
                    break;
                case "Microsoft.ZeroTouch.Administration.ZTServer":
                    if (o != null)
                        xmlElement.InnerText = ((ZTServer)o).Id.ToString();
                    else
                        xmlElement.InnerText = "";
                    break;
                case "Microsoft.ZeroTouch.Administration.ZTDatabase":
                    xmlElement.InnerText = SerializePersistedObject((ZTBasePersistedObject)o);
                    break;
                case "Microsoft.ZeroTouch.Connectivity.Orchestrator.ZTOrchestratorDatabase":
                    xmlElement.InnerText = SerializePersistedObject((ZTBasePersistedObject)o);
                    break;
                case "Microsoft.ZeroTouch.Administration.ZTManagedAccount":
                    if (o != null)
                        xmlElement.InnerText = ((ZTManagedAccount)o).Id.ToString();
                    else
                        xmlElement.InnerText = "";
                    break;
                case "Microsoft.ZeroTouch.Administration.ZTEndpoint":
                    if (o != null)
                        xmlElement.InnerText = ((ZTEndpoint)o).Id.ToString();
                    else
                        xmlElement.InnerText = "";
                    break;
                case "Microsoft.ZeroTouch.ZTSchedule":
                    xmlElement.InnerText = ((ZTSchedule)o).ToString(true);
                    break;
                case "Microsoft.ZeroTouch.Administration.ZTPowerShellScript":
                    if (o != null)
                        xmlElement.InnerText = ((ZTPowerShellScript)o).Id.ToString();
                    else
                        xmlElement.InnerText = "";
                    break;
                case "System.Collections.Hashtable":
                    {
                        XmlElement xmlHT = SerializeCollection(xmlElement, Name, Type, o);
                        xmlElement = xmlHT;
                        break;
                    }
                case "System.Byte[]":
                    {
                        XmlElement xmlHT = SerializeCollection(xmlElement, Name, Type, o);
                        xmlElement = xmlHT;
                        break;
                    }
                case "Microsoft.ZeroTouch.Administration.ZTNotificationEndpointCollection":
                    {
                        XmlElement xmlHT = SerializeCollection(xmlElement, Name, Type, o);
                        xmlElement = xmlHT;
                        break;
                    }
                case "Microsoft.ZeroTouch.Administration.ZTEncryptedString":
                    {
                        if (o != null)
                        {
                            XmlElement xmlHT = SerializeCollection(xmlElement, Name, Type, o);
                            xmlElement = xmlHT;
                        }
                        else
                        {
                            xmlElement.InnerText = "";
                        }
                        break;
                    }
                case "System.Management.Automation.Runspaces.AuthenticationMechanism":
                    xmlElement.InnerText = ((System.Management.Automation.Runspaces.AuthenticationMechanism)o).ToString();
                    break;
                case "Microsoft.ZeroTouch.Administration.ZTBaseActionParameterDefinitionCollection":
                    xmlElement.InnerText = o == null ? "" : ((ZTBaseActionParameterDefinitionCollection)o).Serialize();
                    break;
                case "Microsoft.ZeroTouch.Administration.ZTBaseActionParameterDefinition":
                    xmlElement.InnerText = o == null ? "" : ((ZTBaseActionParameterDefinition)o).Serialize().OuterXml;
                    break;
                case "Microsoft.ZeroTouch.Administration.ZTTicketApprovedConditionCollection":
                    xmlElement.InnerText = o == null ? "" : ((ZTTicketApprovedConditionCollection)o).Serialize();
                    break;
                case "Microsoft.ZeroTouch.Administration.ZTTicketApprovedCondition":
                    xmlElement.InnerText = o == null ? "" : ((ZTTicketApprovedCondition)o).Serialize().OuterXml;
                    break;
                case "Microsoft.ZeroTouch.Administration.ZTKillSwitch":
                    xmlElement.InnerText = o == null ? "" : ((ZTKillSwitch)o).Serialize();
                    break;
                case "Microsoft.ZeroTouch.Administration.ZTHealthCheckResult":
                    xmlElement.InnerText = o == null ? "" : ((ZTHealthCheckResult)o).Serialize();
                    break;
                default:
                    if (o is IZTSerializable)
                    {
                        XmlElement xmlChildElement = ((IZTSerializable)o).Serialize(xmlElement);
                        xmlElement.AppendChild(xmlChildElement);
                    }
                    else if (o is ZTBasePersistedObject)
                    {
                        if (o != null)
                            xmlElement.InnerText = ((ZTBasePersistedObject)o).Id.ToString();
                        else
                            xmlElement.InnerText = "";
                    }
                    else
                    {
                        ZTTraceLogging.LogEvent(ZTLoggingLevel.Warning, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00114", string.Format("Serialization: Object type: {0}. Unsupported property type {1}. Trying to save as string.", this.GetType().AssemblyQualifiedName, Type), "");
                        xmlElement.InnerText = o == null ? string.Empty : o.ToString();
                    }
                    break;
            }

            return xmlElement;
        }

        /// <summary>
        /// Serializes the object as the XmlDocument.
        /// Intended for internal use only.
        /// </summary>
        /// <seealso cref="SerializeField(XmlElement, string, string, string, object)"/>
        /// <seealso cref="EnumFields(Type)"/> 
        /// <owner alias="aldras" />
        public XmlDocument SerializeBasicObject()
        {
            // ZTTraceLogging.LogEvent(ZTLoggingLevel.Information, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00120", string.Format("Serialization started for the object type: {0}.", this.GetType().AssemblyQualifiedName), "");

            Dictionary<string, FieldInfo> fields = EnumFields(this.GetType());
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement xmlRoot = xmlDoc.CreateElement("object");
            xmlRoot.SetAttribute("Type", this.GetType().AssemblyQualifiedName);

            object o;
            XmlElement xmlElement;

            foreach (KeyValuePair<string, FieldInfo> f in fields)
            {
                o = f.Value.GetValue(this);
                xmlElement = SerializeField(xmlRoot, "fld", f.Key, f.Value.FieldType.FullName, o);
                xmlRoot.AppendChild(xmlElement);
            }

            // serialize deprecated properties
            foreach (DictionaryEntry item in _upgradedPersistedProperties)
            {
                XmlElement xmlFld = SerializeField(xmlRoot, "fld", item.Key.ToString(), item.Value.GetType().FullName, item.Value);
                xmlRoot.AppendChild(xmlFld);
            }

            xmlDoc.AppendChild(xmlRoot);

            // ZTTraceLogging.LogEvent(ZTLoggingLevel.Information, ZTLoggingComponent.Core, ZTLoggingCategory.General, "ASO00129", string.Format("Serialization completed for the object type: {0}. {1} properties found.", this.GetType().AssemblyQualifiedName, xmlDoc.FirstChild.ChildNodes.Count.ToString()), "");

            return xmlDoc;
        }
        #endregion Serialization
    }
}
*/