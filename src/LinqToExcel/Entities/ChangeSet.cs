using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LinqToExcel
{
    public class ChangeSet
    {
        private List<ObjectState> trackedList;
        public ChangeSet()
        {
            trackedList = new List<ObjectState>();
        }
        public void AddObject(ObjectState objectState)
        {
            trackedList.Add(objectState);
        }
        public void InsertObject(Object obj)
        {
            foreach (ObjectState os in trackedList)
            {
                if (ObjectState.ReferenceEquals(os.Entity, obj))
                {
                    throw new InvalidOperationException("Object already in list");
                }
            }
            ObjectState osNew = new ObjectState(obj, new List<PropertyManager>());
            osNew.ChangeState = ChangeState.Inserted;
            trackedList.Add(osNew);
        }
        public void DeleteObject(Object obj)
        {
            ObjectState os = (from o in trackedList where Object.ReferenceEquals(o.Entity, obj) == true select o).FirstOrDefault();
            if (os != null)
            {
                if (os.ChangeState == ChangeState.Inserted)
                {
                    trackedList.Remove(os);
                }
                else
                {
                    os.ChangeState = ChangeState.Deleted;
                }
            }
        }
        public List<ObjectState> ChangedObjects
        {
            get { return (from c in trackedList where c.ChangeState != ChangeState.Retrieved select c).ToList(); }
        }
    }
}
