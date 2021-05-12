using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Instrumentation;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;

namespace GroupEditor
{
    class GroupEditor
    {
        private static Schema _schema;
        private static Field _groupNameField, _pinnedField, _locationPointField, _groupIdField, _membersField;

        private Group _group;
        private readonly Document _doc;

        private DataStorage _dataStorage;
        private Entity _storageEntity;

        private readonly string _groupName;
        private readonly bool _pinned;
        private readonly XYZ _locationPoint;
        private readonly ElementId _groupId;
        private IList<ElementId> _members;

        #region Constructors

        public GroupEditor(Group group)
        {
            _doc = group.Document;
            _group = group;
            _groupName = group.GroupType.Name;

            GetDataStorageMembers();
            if (_storageEntity != null)
                throw new ArgumentException(
                    $"Another instance of a the GroupType \"{_groupName}\" is already being edited");

            _pinned = group.Pinned;
            _locationPoint = (group.Location as LocationPoint).Point;
            _groupId = group.Id;
            _members = group.GetMemberIds();
        }

        public GroupEditor(Document doc, string ungroupedGroupName)
        {
            _doc = doc;
            _group = null;
            _groupName = ungroupedGroupName;

            GetDataStorageMembers();
            if (_storageEntity == null)
                throw new InstanceNotFoundException(
                    $"No instance of a the GroupType \"{_groupName}\" is being edited");

            _pinned = _storageEntity.Get<bool>(_pinnedField);
            _locationPoint = _storageEntity.Get<XYZ>(_locationPointField, UnitTypeId.Feet);
            _groupId = _storageEntity.Get<ElementId>(_groupIdField);
            _members = _storageEntity.Get<IList<ElementId>>(_membersField);
        }

        #endregion

        #region Start/finish editing

        public void StartEditing()
        {
            if (_group == null)
                throw new InvalidOperationException("Group already being edited");

            if (_dataStorage != null)
                throw new InvalidOperationException($"Another instance of \"{_groupName}\" is already being edited");

            SetDataStorageFields(_groupName, _pinned, _locationPoint, _groupId, _members);

            var groupType = _group.GroupType;
            var groupTypeMustBeDeleted = groupType.Groups.Size == 1;
            _members = (IList<ElementId>) _group.UngroupMembers();
            if (groupTypeMustBeDeleted)
                _doc.Delete(groupType.Id);

            _group = null;
        }

        public Group FinishEditing()
        {
            if (!GroupElements().Any())
                throw new InvalidOperationException(
                    "None of the elements listed as group members is valid.\n\nPerhaps they have been deleted?");

            _group = _doc.Create.NewGroup(_members);

            // find other instances of the same GroupType and update them
            var oldGroupType = new FilteredElementCollector(_doc)
                .OfClass(typeof(GroupType))
                .Cast<GroupType>()
                .FirstOrDefault(group => group.Name == _groupName);

            if (oldGroupType != null)
            {
                // testing against the group.Id is required because calling both Group.UngroupMembers and
                // GroupType.Groups in the same transaction causes GroupType.Groups to return the ungrouped group.
                // See https://forums.autodesk.com/t5/revit-api-forum/why-does-grouptype-groups-contain-ungrouped-groups/m-p/10292162
                foreach (Group group in oldGroupType.Groups)
                {
                    if (group.Id == _groupId)
                        continue;

                    group.GroupType = _group.GroupType;
                    group.Location.Move((_group.Location as LocationPoint).Point - _locationPoint);
                }

                _doc.Delete(oldGroupType.Id);
            }

            _group.GroupType.Name = _groupName;
            _group.Pinned = _pinned;

            _doc.Delete(_dataStorage.Id);
            _dataStorage = null;
            _storageEntity = null;

            return _group;
        }

        #endregion

        #region Add/remove elements

        public void AddElement(Element element)
        {
            _members.Add(element.Id);
            SetDataStorageFields(members: _members);
        }

        public void AddElement(ElementId elementId)
        {
            _members.Add(elementId);
            SetDataStorageFields(members: _members);
        }

        public void AddElements(IEnumerable<Element> elements)
        {
            foreach (var element in elements)
                _members.Add(element.Id);
            SetDataStorageFields(members: _members);
        }

        public void AddElements(IEnumerable<ElementId> elementsIds)
        {
            foreach (var element in elementsIds)
                _members.Add(element);
            SetDataStorageFields(members: _members);
        }

        public void RemoveElement(Element element)
        {
            _members.Remove(element.Id);
            SetDataStorageFields(members: _members);
        }

        public void RemoveElement(ElementId elementId)
        {
            _members.Remove(elementId);
            SetDataStorageFields(members: _members);
        }

        #endregion

        #region Collectors

        public static List<string> GetNamesOfGroupsBeingEdited(Document doc)
        {
            GetSchema();

            var groupEditorDataStorages = new FilteredElementCollector(doc).OfClass(typeof(DataStorage));

            return groupEditorDataStorages.Select(element => element.GetEntity(_schema))
                .Where(dataStorage => dataStorage.IsValid())
                .Select(dataStorage => dataStorage.Get<string>("GroupName"))
                .OrderBy(name => name)
                .ToList();
        }

        public IEnumerable<Element> GroupElements()
        {
            return _members
                .Select(elementId => _doc.GetElement(elementId))
                .Where(element => element != null)
                .ToList();
        }

        #endregion

        #region Schema and DataStorage management

        private static void GetSchema()
        {
            if (_schema != null)
                return;

            var schemaGuid = new Guid("98140745-875d-4e4e-8e5e-a36146a4e847");
            _schema = Schema.Lookup(schemaGuid);
            if (_schema == null)
            {
                var schemaBuilder = new SchemaBuilder(schemaGuid);
                schemaBuilder.SetSchemaName("GroupEditor");
                schemaBuilder.SetReadAccessLevel(AccessLevel.Public);
                schemaBuilder.SetWriteAccessLevel(AccessLevel.Public);
                schemaBuilder.AddSimpleField("GroupName", typeof(string));
                schemaBuilder.AddSimpleField("Pinned", typeof(bool));
                schemaBuilder.AddSimpleField("LocationPoint", typeof(XYZ)).SetSpec(SpecTypeId.Length);
                schemaBuilder.AddSimpleField("GroupId", typeof(ElementId));
                schemaBuilder.AddArrayField("Members", typeof(ElementId));
                _schema = schemaBuilder.Finish();
            }

            _groupNameField = _schema.GetField("GroupName");
            _pinnedField = _schema.GetField("Pinned");
            _locationPointField = _schema.GetField("LocationPoint");
            _groupIdField = _schema.GetField("GroupId");
            _membersField = _schema.GetField("Members");
        }

        private void GetDataStorageMembers()
        {
            GetSchema();

            var dataStorages = new FilteredElementCollector(_doc).OfClass(typeof(DataStorage));
            foreach (var element in dataStorages)
            {
                var entity = element.GetEntity(_schema);
                if (entity.IsValid() && entity.Get<string>(_groupNameField) == _groupName)
                {
                    _storageEntity = entity;
                    _dataStorage = element as DataStorage;
                    return;
                }
            }
        }

        private void SetDataStorageFields(string groupName = null, bool? pinned = null, XYZ localPoint = null,
            ElementId groupId = null, IList<ElementId> members = null)
        {
            if (_dataStorage == null)
            {
                _dataStorage = DataStorage.Create(_doc);
                _storageEntity = new Entity(_schema);
            }

            if (groupName != null) _storageEntity.Set("GroupName", groupName);
            if (pinned != null) _storageEntity.Set("Pinned", (bool) pinned);
            if (localPoint != null) _storageEntity.Set("LocationPoint", localPoint, UnitTypeId.Feet);
            if (groupId != null) _storageEntity.Set("GroupId", groupId);
            if (members != null) _storageEntity.Set("Members", members);

            _dataStorage.SetEntity(_storageEntity);
        }

        public static void DeleteDataStorageSchemaEntity(Document doc, string groupName)
        {
            GetSchema();

            var dataStorages = new FilteredElementCollector(doc).OfClass(typeof(DataStorage));
            foreach (var element in dataStorages)
            {
                var existingDataStorage = element.GetEntity(_schema);
                if (existingDataStorage.IsValid() && existingDataStorage.Get<string>(_groupNameField) == groupName)
                {
                    doc.Delete(element.Id);

                    var groupType = new FilteredElementCollector(doc)
                        .OfClass(typeof(GroupType))
                        .FirstOrDefault(gt => gt.Name == groupName) as GroupType;

                    if (groupType != null && groupType.Groups.Size == 0)
                        doc.Delete(groupType.Id);

                    return;
                }
            }

            throw new InstanceNotFoundException($"DataStorage for group \"{groupName}\" not found");
        }

        #endregion
    }
}