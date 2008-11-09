using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel
{
    public class ObjectState
    {
        private List<PropertyManager> propList;
        private object entity;
        private ChangeState state;
        public ObjectState(object entity, List<PropertyManager> props)
        {
            this.entity = entity;
            this.propList = props;
            state = ChangeState.Retrieved;
            if (entity is System.ComponentModel.INotifyPropertyChanged)
            {
                System.ComponentModel.INotifyPropertyChanged i = (System.ComponentModel.INotifyPropertyChanged)entity;
                i.PropertyChanged += new System.ComponentModel.PropertyChangedEventHandler(i_PropertyChanged);
            }
        }
        public List<PropertyManager> Properties
        {
            get { return this.propList; }
        }
        public PropertyManager GetProperty(string propertyName)
        {
            return (from p in propList where p.PropertyName == propertyName select p).FirstOrDefault();
        }
        public List<PropertyManager> ChangedProperties
        {
            get { return (from p in propList where p.HasChanged == true select p).ToList(); }
        }
        public ChangeState ChangeState
        {
            get { return state; }
            set { state = value; }
        }
        public Object Entity
        {
            get { return this.entity; }
        }
        public void i_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            PropertyManager pm = (from p in propList where p.HasChanged == false && p.PropertyName == e.PropertyName select p).FirstOrDefault();
            if (pm != null)
            {
                pm.HasChanged = true;
                if (state == ChangeState.Retrieved)
                    state = ChangeState.Updated;
            }
        }
    }
}
