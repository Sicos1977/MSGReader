namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf attribute
    /// </summary>
    internal class Attribute
    {
        #region Properties
        /// <summary>
        /// attribute's name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// value
        /// </summary>
        public int Value { get; set; }
        #endregion

        #region Constructor
        /// <summary>
        /// Initialize instance
        /// </summary>
        public Attribute()
        {
            Value = int.MinValue;
            Name = null;
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            return Name + "=" + Value;
        }
        #endregion
    }

    /// <summary>
    /// RTF attribute list
    /// </summary>
    internal class AttributeList : System.Collections.CollectionBase
    {
        #region GetItem
        public Attribute GetItem(int index)
        {
            return (Attribute) List[index];
        }
        #endregion

        #region ByName
        public int this[string name]
        {
            get
            {
                foreach (Attribute attribute in this)
                {
                    if (attribute.Name == name)
                        return attribute.Value;
                }
                return int.MinValue;
            }
            set
            {
                foreach (Attribute attribute in this)
                {
                    if (attribute.Name == name)
                    {
                        attribute.Value = value;
                        return;
                    }
                }
                
                var attr = new Attribute {Name = name, Value = value};
                List.Add(attr);
            }
        }
        #endregion

        #region Add
        public int Add(Attribute item)
        {
            return List.Add(item);
        }

        public int Add(string name, int value)
        {
            var attribute = new Attribute {Name = name, Value = value};
            return List.Add(attribute);
        }
        #endregion

        #region Remove
        public void Remove(Attribute attribute)
        {
            List.Remove(attribute);
        }

        public void Remove(string name)
        {
            for (var count = Count - 1; count >= 0; count--)
            {
                var attribute = (Attribute) List[count];
                if (attribute.Name == name)
                    List.RemoveAt(count);
            }
        }
        #endregion

        #region Contains
        public bool Contains(Attribute attribute)
        {
            return List.Contains(attribute);
        }

        public bool Contains(string name)
        {
            foreach (Attribute attribute in this)
            {
                if (attribute.Name == name)
                    return true;
            }
            return false;
        }
        #endregion

        #region Clone
        /// <summary>
        /// Clone this object
        /// </summary>
        /// <returns></returns>
        public AttributeList Clone()
        {
            var attributeList = new AttributeList();
            foreach (Attribute item in this)
            {
                var newItem = new Attribute {Name = item.Name, Value = item.Value};
                attributeList.List.Add(newItem);
            }
            return attributeList;
        }
        #endregion
    }
}
