using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NHibernate.UserTypes;
using NHibernate;
using NHibernate.SqlTypes;

namespace Russell.RADAR.POC.Entities.Content
{
    /// <summary>
    /// Allow NHibernate persistence...
    /// </summary>
    public class FormattedContentUserType : IUserType
    {
        public object Assemble(object cached, object owner)
        {
            return cached;
        }

        public object DeepCopy(object value)
        {
            if (value == null)
                return null;

            return ((FormattedContent)value).Clone();
        }

        public object Disassemble(object value)
        {
            return value;
        }

        public int GetHashCode(object x)
        {
            return x.GetHashCode();
        }

        public bool IsMutable
        {
            get { return true; }
        }

        public object NullSafeGet(System.Data.IDataReader rs, string[] names, object owner)
        {
            var xhtmlRepresentation = (string)NHibernateUtil.String.NullSafeGet(rs, names);
            var result = new FormattedContent();
            result.FromXHTML(xhtmlRepresentation);
            return result;
        }

        public void NullSafeSet(System.Data.IDbCommand cmd, object value, int index)
        {
            NHibernateUtil.String.NullSafeSet(cmd, ((FormattedContent)value).ToXHTML(), index);
        }

        public object Replace(object original, object target, object owner)
        {
            return original;
        }

        public Type ReturnedType
        {
            get { return typeof(FormattedContent); }
        }

        public NHibernate.SqlTypes.SqlType[] SqlTypes
        {
            get { return new[] { SqlTypeFactory.GetString(10000) }; }
        }

        bool IUserType.Equals(object x, object y)
        {
            return ((FormattedContent)x).ToXHTML().Equals(((FormattedContent)y).ToXHTML());
        } 
    }
}
