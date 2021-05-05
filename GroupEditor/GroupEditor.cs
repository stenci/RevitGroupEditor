using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using InvalidOperationException = System.InvalidOperationException;

namespace GroupEditor
{
    class GroupEditor
    {
        private Group _group;
        private readonly Document _doc;

        private readonly string _groupName;
        private readonly bool _pinned;
        private readonly XYZ _locationPoint;

        private static Schema _schema;
        private static Field _fieldGroupName, _fieldPinned, _fieldLocationPoint;

        public GroupEditor(Group group)
        {
            GetSchema();

            _doc = group.Document;
            _group = group;

            _groupName = group.GroupType.Name;
            _pinned = group.Pinned;
            _locationPoint = (group.Location as LocationPoint).Point;
        }

        public GroupEditor(Document doc, string ungroupedGroupName)
        {
            GetSchema();

            _doc = doc;
            _group = null;
            _groupName = ungroupedGroupName;
            var firstElementSchema = GroupElements().First().GetEntity(_schema);
            _pinned = firstElementSchema.Get<bool>("Pinned");
            _locationPoint = firstElementSchema.Get<XYZ>("LocationPoint", UnitTypeId.Feet);
        }

        public void StartEditing()
        {
            if (_group == null)
                throw new InvalidOperationException("The group is already open");

            foreach (var id in _group.UngroupMembers())
                SetElementSchema(_doc.GetElement(id));

            _group = null;
        }

        public void AddElement(Element element)
        {
            SetElementSchema(element);
        }

        private void SetElementSchema(Element element)
        {
            var schemaEntity = new Entity(_schema);
            schemaEntity.Set(_fieldGroupName, _groupName);
            schemaEntity.Set(_fieldPinned, _pinned);
            schemaEntity.Set(_fieldLocationPoint, _locationPoint, UnitTypeId.Feet);
            element.SetEntity(schemaEntity);
        }

        public void FinishEditing()
        {
            var elements = new List<ElementId>();
            foreach (var element in GroupElements())
            {
                elements.Add(element.Id);
                element.DeleteEntity(_schema);
            }

            _group = _doc.Create.NewGroup(elements);

            var oldGroupType = new FilteredElementCollector(_doc)
                .OfClass(typeof(GroupType))
                .Cast<GroupType>()
                .FirstOrDefault(group => group.Name == _groupName);

            if (oldGroupType != null)
            {
                foreach (Group group in oldGroupType.Groups)
                {
                    group.GroupType = _group.GroupType;
                    group.Location.Move((_group.Location as LocationPoint).Point - _locationPoint);
                }

                _doc.Delete(oldGroupType.Id);
            }

            _group.GroupType.Name = _groupName;
            _group.Pinned = _pinned;
        }

        public static IEnumerable<string> GetNamesOfGroupsBeingEdited(Document doc)
        {
            GetSchema();

            return new FilteredElementCollector(doc)
                .WherePasses(new ExtensibleStorageFilter(_schema.GUID))
                .Select(element => element.GetEntity(_schema).Get<string>("GroupName"))
                .Distinct();
        }

        public void DeleteAllGroupEditorSchemas()
        {
            foreach (var element in GroupElements())
                element.DeleteEntity(_schema);
        }

        public IEnumerable<Element> GroupElements()
        {
            return new FilteredElementCollector(_doc)
                .WherePasses(new ExtensibleStorageFilter(_schema.GUID))
                .Where(element => element.GetEntity(_schema).Get<string>("GroupName") == _groupName);
        }

        private static void GetSchema()
        {
            if (_schema != null) goto getSchemaFields;

            var schemaGuid = new Guid("98140745-875d-4e4e-8e5e-a36146a4e844");
            _schema = Schema.Lookup(schemaGuid);
            if (_schema != null) goto getSchemaFields;

            var schemaBuilder = new SchemaBuilder(schemaGuid);
            schemaBuilder.SetSchemaName("GroupEditor");
            schemaBuilder.SetReadAccessLevel(AccessLevel.Public);
            schemaBuilder.SetWriteAccessLevel(AccessLevel.Public);
            schemaBuilder.AddSimpleField("GroupName", typeof(string));
            schemaBuilder.AddSimpleField("Pinned", typeof(bool));
            schemaBuilder.AddSimpleField("LocationPoint", typeof(XYZ)).SetSpec(SpecTypeId.Length);
            _schema = schemaBuilder.Finish();

            getSchemaFields:
            if (_fieldGroupName != null) return;
            _fieldGroupName = _schema.GetField("GroupName");
            _fieldPinned = _schema.GetField("Pinned");
            _fieldLocationPoint = _schema.GetField("LocationPoint");
        }
    }
}