namespace MsgReader.Rtf
{
    internal class StringAttribute
    {
        #region Properties
        /// <summary>
        /// Name of the attribute
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Value of the attribute
        /// </summary>
        public string Value { get; set; }
        #endregion

        #region Constructor
        public StringAttribute()
        {
            Value = null;
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

    internal class StringAttributeCollection : System.Collections.CollectionBase
    {
        #region Properties
        public string this[string name]
        {
            get
            {
                foreach (StringAttribute stringAttribute in this)
                {
                    if (stringAttribute.Name == name)
                        return stringAttribute.Value;
                }
                return null;
            }
            set
            {
                foreach (StringAttribute stringAttribute in this)
                {
                    if (stringAttribute.Name == name)
                    {
                        if (value == null)
                            List.Remove(stringAttribute);
                        else
                            stringAttribute.Value = value;
                        return;
                    }
                }
                if (value != null)
                {
                    var stringAttribute = new StringAttribute { Name = name, Value = value };
                    List.Add(stringAttribute);
                }
            }
        }
        #endregion

        #region Add
        /// <summary>
        /// Add's an item to the list and returns the index
        /// </summary>
        /// <param name="stringAttribute"></param>
        /// <returns></returns>
        public int Add(StringAttribute stringAttribute)
        {
            return List.Add(stringAttribute);
        }
        #endregion

        #region Remove
        /// <summary>
        /// Removes an item from the list
        /// </summary>
        /// <param name="stringAttribute"></param>
        public void Remove(StringAttribute stringAttribute)
        {
            List.Remove(stringAttribute);
        }
        #endregion
    }
}
