# LogFileViewer.VB6
Log file viewer -- written in VB6

# Enhancement Ideas

*	Don’t allow log format match header to be empty.

*	Open source code file associated with error.  This may be VERY difficult to implement.

*	Export to real CSV format in addition to tab-delimited.

*	Enable writing ASCII files in addition to Unicode.

*	Improve find by highlighting the found text in either the grid or the detail view or maybe both.

*	Allow option to show item details as two columns (list view control) instead of just via a rich text control.

*	Allow top-down sequential order in addition to bottom-up.

*	Support list of favorite log files – adding, removing, editing, and opening.

*	Save column widths when no format is active.

# Completed Changes 

*Insert new sections here ... with title: for Version X or in [time-period] or on [date]*

## Changes Completed for Version 8.0.0.47 (March 2006)

•	Limit the time that the application is busy updating the display to no more than 10%.  The time to update the display with file changes is measured.  Then, this value plus the duration of the previous 7 updates are averaged.  The interval is set to the average times 10 or to one second – whichever is more.  This change has basically no effect on monitoring a file that is on fast storage such as a local disk.  But, this change improves the user experience for a file that is on a slow device such as a network drive.  Before, the application might spend most of its time updating the display for a file on a slow network drive – not allowing the user to navigate the display or select menu items.  Now, the application pauses longer between updates for such a file – allowing the user to control the application.

•	Improve feedback on errors during file load/update.

•	Allow the user to specify a non-existent path for monitoring via the Open command.  The app already allowed monitoring a non-existent path for the startup file and via the Open Recent command, but the app prevented selection of a non-existent path via the Open command.

•	Insure that stream view control (RichTextBox) maintains focus after update.

## Changes Completed for Version 8.0.0.2 (February 2006)

•	Modified major version number to 8 (from 0) to indicate that the app is ready for general use.

## Changes Completed for Version 0.0.0.55 (October 2005)

•	Improve support for other log file formats.  The tool now supports user-defined log formats.  The user can 
define column formats where each line field is delimited with the same character.  Column formats may or may
not have a header line.  Also, non-column formats (1 column) are supported.  The column widths set by the user
are saved for each format.

•	MRU file list.

•	Improve dialog that is used when file is larger than recommended size.  Use a custom form instead of a message box and buttons that have better captions than ‘yes’, ‘no’ and ‘cancel’.

•	De-reference shell link files on open.  This was being done automatically by the standard file dialog control if the file was opened via the Open menu item.  But, this was not being done if the file was opened from a startup command-line parameter or a run-time drag-and-drop operation.

## Changes Completed for Version 7.3.0.40 (September 2004)

•	Disable updating during merge-multiple operation.

•	Enable reading ASCII files in addition to Unicode.

•	Disable Copy menu item if nothing selected.  Not only is this a better user-interface, but this also fixes an error that was occurring if the user selected Copy when there was no file loaded.

•	Disable Select All, Find, and Find Next menu items if no log lines loaded.

## Changes Completed for Version 7.3.0.? (May 2004)

•	Add add-to-log tool for adding new entries to the log file via the epLogger object.

•	Support integration of multiple files with possibly overlapping time-frames.  The major use case for this is analyzing log files from multiple machines since errors may propagate from one machine to another.  This capability is implemented as the new merge-multiple-files tool.

•	Fix bug: opening empty file causes file read error.  The file-system-object fails to read a Unicode file’s contents if it is empty.  Need to handle the empty file condition better.

•	Change SaveAs to copy the active file instead of just saving all of the entries in the list-view.  The list-view may not contain all of the items of the file due to roll-up.  It might be confusing to the user if a SaveAs only saves the list-view items instead of the entire file contents.  Note: the user could save all of the entries of the list-view by doing SelectAll and then Export.  

•	Since the tab-delimited format (what the logger generates) seems much more useful than the XML format (no tools read it), use the tab-delimited format for an output drag-and-drop operation.  Also, make the tab-delimited format the default for the Export function.

•	Prevent a drop operation when drag source is the same application.  The old behavior of allowing this lead to undesirable behavior of clearing the display if the user accidentally clicked-held-and-moved the mouse in the list-view.

•	Support clipboard text format for outgoing drag-and-drop of log entries in addition to file format.  This allows the user to drag log entries into a text editor that supports a text format drop operation.  For example, Word and Excel support text drop, but Notepad does not.

•	Support copy-to-clipboard (text format) for log entries.  This allows the user to copy log entries to the clipboard and then paste them into any text editor.  The RTF control that shows item details comes with support for Ctrl-C to copy-to-clipboard. But, there is no such built-in support in the log entries list-view control.  Add a Copy (Ctrl-C) menu item under Edit main menu to support copy-to-clipboard.  Insure that this new menu item as well as Ctrl-C short cut work equally well with the log view as well as the details view.  The details view will support both plain text as well as RTF, although the log view will support only plain text.

## Changes Completed for Older Versions

•	Support drag-and-drop from Windows Explorer in order to open a file.

•	Add find capability.

•	Show file-line as # column instead of item number.

•	Maintain selected log item when refresh and (un)select roll-up.

•	Fix bug: Error when rollup is on and file has numeric ID value.

•	Handle files with no header row in file.  Hard-code a fixed set of column headers to use in the case where no header line is found.  Downside: If the file format changes, the log viewer will not handle a file with not headers very well.

•	Improve efficiency by loading less than the entire file on each refresh loop.

•	Design and use a new application icon instead of using the one from the DebugToolsViewer application.
