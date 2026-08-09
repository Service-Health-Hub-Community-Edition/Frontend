namespace Microsoft.ServiceHealthHub.Core
{
    public class EntityProperty
    {
        public string entityProperty { get; set; }
        public string displayName { get; set; }
        public string type { get; set; }
        public bool hidden { get; set; }
    }

    public class Component
    {
        MSSHNotificationDatabase _db = new MSSHNotificationDatabase();

        Guid id;
        string name;
        string internalName;
        string icon;
        List<string> capabilities;
        List<EntityProperty> entityProperties;

        public Guid Id { get { return id; } set { id = value; } }
        public string Name { get { return name; } set { name = value; } }
        public string InternalName { get { return internalName; } set { internalName = value; } }
        public string Icon { get { return icon; } set { icon = value; } }
        public List<string> Capabilities { get { return capabilities; } }
        public List<EntityProperty> EntityProperties { get { return entityProperties; } }

        Component()
        {
            capabilities = new List<string>();
            entityProperties = new List<EntityProperty>();
        }

        Component(Guid id)
        {
            List<MSComponent> dbComponents = _db.GetComponent(id);
            if (dbComponents != null && dbComponents.Count >= 1)
            {

            }
        }
    }
}
