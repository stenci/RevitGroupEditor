using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace GroupEditor
{
    [Transaction(TransactionMode.Manual)]
    public class GroupEditorStartEditing : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var app = commandData.Application;
            var uidoc = app.ActiveUIDocument;
            var doc = uidoc.Document;

            var selectedElements = uidoc.Selection.GetElementIds();
            if (selectedElements.Count != 1)
            {
                TaskDialog.Show("Group Editor",
                    "Please select a group and try again.");
                return Result.Succeeded;
            }

            var group = doc.GetElement(selectedElements.First()) as Group;
            if (group == null)
            {
                TaskDialog.Show("Group Editor",
                    "Please select a group and try again.");
                return Result.Succeeded;
            }

            using (var tx = new Transaction(doc))
            {
                tx.Start("Group Editor Start Editing");

                var groupEditor = new GroupEditor(group);
                groupEditor.StartEditing();

                tx.Commit();
            }

            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class GroupEditorFinishEditing : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var app = commandData.Application;
            var uidoc = app.ActiveUIDocument;
            var doc = uidoc.Document;

            var groupName = Utils.PickOneOfTheGroupsBeingEdited(doc,
                "Select the name of the group of the elements to regroup.");

            if (groupName == null)
                return Result.Cancelled;

            using (var tx = new Transaction(doc))
            {
                tx.Start("Group Editor Finish Editing");

                var groupEditor = new GroupEditor(doc, groupName);
                groupEditor.FinishEditing();

                tx.Commit();
            }

            TaskDialog.Show("Group Editor",
                $"The elements of the group \"{groupName}\" have been regrouped.");

            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class GroupEditorAddToGroup : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var app = commandData.Application;
            var uidoc = app.ActiveUIDocument;
            var doc = uidoc.Document;

            var selectedElements = uidoc.Selection.GetElementIds();
            if (selectedElements.Count == 0)
            {
                TaskDialog.Show("Group Editor",
                    "Please select the elements to add to the group and try again.");
                return Result.Succeeded;
            }

            var groupName = Utils.PickOneOfTheGroupsBeingEdited(doc,
                "Select the name of the group you want to add the pre-selected elements to.");

            if (groupName == null)
                return Result.Cancelled;

            using (var tx = new Transaction(doc))
            {
                tx.Start("Group Editor Add To Group");

                var groupEditor = new GroupEditor(doc, groupName);

                foreach (var elementId in selectedElements)
                    groupEditor.AddElement(doc.GetElement(elementId));

                tx.Commit();
            }

            TaskDialog.Show("Group Editor",
                $"The pre-selected elements have been added to group \"{groupName}\".");

            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class GroupEditorPurge : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var app = commandData.Application;
            var uidoc = app.ActiveUIDocument;
            var doc = uidoc.Document;

            var groupName = Utils.PickOneOfTheGroupsBeingEdited(doc,
                "Select the name of the group whose entities' group info you want to remove.");

            if (groupName == null)
                return Result.Cancelled;

            using (var tx = new Transaction(doc))
            {
                tx.Start("Group Editor Purge");

                var groupEditor = new GroupEditor(doc, groupName);
                groupEditor.DeleteEntitySchemas();

                tx.Commit();
            }

            TaskDialog.Show("Group Editor",
                $"Group \"{groupName}\" has been removed from all entities' info.");

            return Result.Succeeded;
        }
    }

    static class Utils
    {
        public static string PickOneOfTheGroupsBeingEdited(Document doc, string mainInstruction)
        {
            var groupNames = GroupEditor.GetNamesOfGroupsBeingEdited(doc).ToList();
            groupNames.Sort();

            if (groupNames.Count == 0)
            {
                TaskDialog.Show("Group Editor",
                    "There are no groups being edited.");
                return null;
            }

            if (groupNames.Count == 1)
                return groupNames[0];

            var dialog = new TaskDialog("Group Editor") {MainInstruction = mainInstruction};

            if (groupNames.Count > 4)
                dialog.MainContent = $"Only 4 of the {groupNames.Count} groups being edited are listed.";

            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, groupNames[0]);
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, groupNames[1]);
            if (groupNames.Count >= 3) dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, groupNames[2]);
            if (groupNames.Count >= 4) dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink4, groupNames[3]);

            var result = dialog.Show();

            switch (result)
            {
                case TaskDialogResult.Cancel: return null;
                case TaskDialogResult.CommandLink1: return groupNames[0];
                case TaskDialogResult.CommandLink2: return groupNames[1];
                case TaskDialogResult.CommandLink3: return groupNames[2];
                case TaskDialogResult.CommandLink4: return groupNames[3];
                default: throw new ArgumentOutOfRangeException();
            }
        }
    }
}