While waiting for [this idea](https://forums.autodesk.com/t5/revit-ideas/group-edit-mode-in-api/idc-p/10293127)
to be addressed, I use this `GroupEditor` class to manage the ungroup - change - regroup process.

The `GroupEditor` allows to ungroup one instance of one `GroupType`, so the user or the APIs can
modify the group member entities, delete or add some. When the changes are finished it is
possible to regroup all the entities into one new group and assign the new `GroupType` to all
the other instances of the old `GroupType`.

Use `GroupEditor.StartEditing()` to ungroup the group and create a `DataStorage` containing 
the list of group members and other group properties. 

Use `GroupEditor.AddElements()` and `GroupEditor.RemoveElements()` to add and remove some
elements from the list in the `DataStorage` object. 

Use `GroupEditor.FinishEditing()` to regroup all the elements. If more instances exist for the 
same `GroupType`, they all are updated to match the newly generated group.

After using `GroupEditor.StartEditing()`, it is possible to give the control back to the user
interface and call `GroupEditor.FinishEditing()` later, even after saving, closing and
reopening the project.

Unfortunately tons of the group properties are lost, like reference level, excluded entities,
some constraints (not all), etc. Some of the properties are lost because Revit APIs do not
manage them, some are lost because I didn't need to manage them yet.
