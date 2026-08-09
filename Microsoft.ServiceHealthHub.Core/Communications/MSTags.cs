using Microsoft.Azure.Amqp.Framing;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Microsoft.ServiceHealthHub.Core
{
    public class MSTagDefinition
    {
        MSSHNotificationDatabase _db = new MSSHNotificationDatabase();

        int m_id = -1;
        Guid m_tagId = new Guid();
        string m_name = "";
        string? m_type;
        int m_itemCount = 0;
        DateTime? m_lastUsed;

        public int Id { get { return m_id; } }
        public Guid TagId { get { return m_tagId; } }
        public string Name { 
            get { return m_name; }
            set {
                if (string.IsNullOrWhiteSpace(value))
                    throw new ArgumentException("Name property cannot be empty");
                if (value.Length < 2 || value.Length > 64)
                    throw new ArgumentException("Name must be 2-64 characters string");
                m_name = value;
            }
        
        }
        public string? Type { get { return m_type; } }
        public int ItemCount { get { return m_itemCount; } }
        public DateTime? LastUsed {  get { return m_lastUsed; } }

        public MSTagDefinition() { }

        public MSTagDefinition(Guid tagId)
        {
            List<MSPropertyBag> res = _db.GetTagDefinition(tagId);

            if (res.Count > 0)
            {
                MSPropertyBag data = res[0];
                m_id = data["Id"] is DBNull ? -1 : (int)data["Id"];
                m_tagId = data["TagId"] is DBNull ? Guid.Empty : (Guid)data["TagId"];
                m_name = data["Name"] is DBNull ? string.Empty : (string)data["Name"];
                m_type = data["Type"] is DBNull ? null : (string)data["Type"];
                m_itemCount = data["ItemCount"] is DBNull ? 0 : (int)data["ItemCount"];
                m_lastUsed = data["LastUsed"] is DBNull ? null : (DateTime)data["LastUsed"];
            }
        }

        public MSTagDefinition(string name, string targetGroup)
        {
            m_name = name;
            m_type = targetGroup;
            m_tagId = Guid.NewGuid();

            Update();
        }

        public void Move(string? targetGroup)
        {
            if (m_id > -1)
                _db.MoveTagDefinition(m_tagId, targetGroup);
        }

        public void Update()
        {
            _db.AddTagDefinition(m_tagId, m_name, m_type);

            if (m_id == -1)
            {
                List<MSPropertyBag> res = _db.GetTagDefinition(m_tagId);

                if (res.Count > 0)
                {
                    MSPropertyBag data = res[0];
                    m_id = data["Id"] is DBNull ? -1 : (int)data["Id"];
                    m_itemCount = data["ItemCount"] is DBNull ? 0 : (int)data["ItemCount"];
                    m_lastUsed = data["LastUsed"] is DBNull ? null : (DateTime)data["LastUsed"];
                }
            }
            
        }

        public void Delete()
        {
            if (m_id > -1)
            {
                _db.RemoveTagDefinition(m_tagId);

                m_id = -1;
                m_itemCount = 0;
                m_lastUsed = null;
                m_name = "";
                m_type = null;
            }
        }

        public static MSTagDefinition CreateInstance(MSPropertyBag data)
        {
            MSTagDefinition instance = new MSTagDefinition()
            {
                m_id = data["Id"] is DBNull ? -1 : (int)data["Id"],
                m_tagId = data["TagId"] is DBNull ? Guid.Empty : (Guid)data["TagId"],
                m_name = data["Name"] is DBNull ? string.Empty : (string)data["Name"],
                m_type = data["Type"] is DBNull ? null : (string)data["Type"],
                m_itemCount = data["ItemCount"] is DBNull ? 0 : (int)data["ItemCount"],
                m_lastUsed = data["LastUsed"] is DBNull ? null : (DateTime)data["LastUsed"]
            };

            return instance;
        }
    }

    public class MSTagDefinitions: List<MSTagDefinition>
    {
        public MSTagDefinitions() { }

        public static MSTagDefinitions CreateInstance() 
        {
            MSTagDefinitions instance = new MSTagDefinitions();
            MSSHNotificationDatabase database = new MSSHNotificationDatabase();
            List<MSPropertyBag> data = database.GetTagDefinitions();
            foreach (MSPropertyBag tagDef in data)
            {
                MSTagDefinition tagDefinition = MSTagDefinition.CreateInstance(tagDef);
                instance.Add(tagDefinition);
            }
            return instance;
        }
    }

    public class MSTag
    {
        MSSHNotificationDatabase _db = new MSSHNotificationDatabase();
        int m_id = -1;
        string m_messageId = "";
        string m_type = "";
        Guid m_tagId = Guid.Empty;
        DateTime? m_modified;

        public int Id { get { return m_id; } }
        public string MessageId
        {
            get { return m_messageId; }
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                    throw new ArgumentException("MessageId property cannot be empty");
                if (value.Length < 2 || value.Length > 64)
                    throw new ArgumentException("MessageId must be 2-64 characters string");
                m_messageId = value;
            }

        }

        public string Type
        {
            get { return m_type; }
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                    throw new ArgumentException("Type property cannot be empty");
                if (value.Length < 2 || value.Length > 32)
                    throw new ArgumentException("Type must be 2-32 characters string");
                m_type = value;
            }

        }

        public Guid TagId
        {
            get { return m_tagId; }
            set {  m_tagId = value; }
        }
        public DateTime? Modified { get { return m_modified; } }

        public MSTag() { }

        public MSTag(string messageId, string type, Guid tagId)
        {
            m_messageId = messageId;
            m_type = type;
            m_tagId = tagId;
        }

        public void Update()
        {
            _db.AddTag(m_messageId, m_type, m_tagId);

            if (m_id == -1)
            {
                List<MSPropertyBag> res = _db.GetTag(m_messageId, m_type, m_tagId);

                if (res.Count > 0)
                {
                    MSPropertyBag data = res[0];
                    m_id = data["Id"] is DBNull ? -1 : (int)data["Id"];
                    m_modified = data["Modified"] is DBNull ? null : (DateTime)data["Modified"];
                }
            }
        }

        public void Delete()
        {
            if (m_id > -1)
            {
                _db.RemoveTag(m_messageId, m_type, m_tagId);

                m_id = -1;
                m_messageId = "";
                m_type = "";
                m_modified = null;
                m_tagId = Guid.Empty;
            }
        }

        public static MSTag? CreateInstance(string messageId, string type, Guid tagId)
        {
            MSTag? instance = null;
            MSSHNotificationDatabase db = new MSSHNotificationDatabase();
            List<MSPropertyBag> res = db.GetTag(messageId, type, tagId);
            if (res.Count > 0)
            {
                MSPropertyBag data = res[0];
                instance = new()
                {
                    m_id = data["Id"] is DBNull ? -1 : (int)data["Id"],
                    m_messageId = data["MessageId"] is DBNull ? string.Empty : (string)data["MessageId"],
                    m_type = data["Type"] is DBNull ? string.Empty : (string)data["Type"],
                    m_tagId = data["TagId"] is DBNull ? Guid.Empty : (Guid)data["TagId"],
                    m_modified = data["Modified"] is DBNull ? null : (DateTime)data["Modified"]
                };
            }

            return instance;
        }

        public static MSTag CreateInstance(MSPropertyBag data)
        {
            MSTag instance = new()
            {
                m_id = data["Id"] is DBNull ? -1 : (int)data["Id"],
                m_messageId = data["MessageId"] is DBNull ? string.Empty : (string)data["MessageId"],
                m_type = data["Type"] is DBNull ? string.Empty : (string)data["Type"],
                m_tagId = data["TagId"] is DBNull ? Guid.Empty : (Guid)data["TagId"],
                m_modified = data["Modified"] is DBNull ? null : (DateTime)data["Modified"]
            };

            return instance;
        }
    }

    public class MSTags : List<MSTag>
    {
        public MSTags() { }

        public static MSTags CreateInstance(string messageId, string type)
        {
            MSTags instance = new MSTags();
            MSSHNotificationDatabase database = new MSSHNotificationDatabase();
            List<MSPropertyBag> data = database.GetTags(messageId, type);
            foreach (MSPropertyBag tag in data)
            {
                MSTag tagObj = MSTag.CreateInstance(tag);
                instance.Add(tagObj);
            }
            return instance;
        }
    }

    public class MSCommunicationTag
    {
        MSSHNotificationDatabase _db = new MSSHNotificationDatabase();
        string m_messageId = "";
        string? m_tag = "";
        Guid? m_tagId = Guid.Empty;

        public string MessageId { get { return m_messageId; } }
        public string? Tag { get { return m_tag; } }
        public Guid? TagId { get { return m_tagId; } }

        public MSCommunicationTag() { }

        public MSCommunicationTag(string messageId, string? tag, Guid? tagId)
        {
            m_messageId = messageId;
            m_tag = tag;
            m_tagId = tagId;
        }

        public static MSCommunicationTag CreateInstance(MSPropertyBag data)
        {
            MSCommunicationTag instance = new()
            {
                m_messageId = data["MessageId"] is DBNull ? string.Empty : (string)data["MessageId"],
                m_tag = data["Tag"] is DBNull ? null : (string)data["Tag"],
                m_tagId = data["TagId"] is DBNull ? null : (Guid)data["TagId"]
            };

            return instance;
        }
    }
}
