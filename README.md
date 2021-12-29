# Send to Obsidian

This is an Outlook addin (read: plugin) that sends the currently opened/selected emailchain as a markdown note to Obsidian.

The idea for this addin is courtesy of Discord user `Namtrah#6370` who wanted this functionality!

## Issues

* There's no way to use the OfficeJS API to grab attachments. 
* Authentication using axios or Microsoft's own library fails for some reason.
* Does not actually export anything.
* Assets need to be replaced.
* Taskpane needs to be designed (or removed in case we simply want a single button to export, and don't need settings)