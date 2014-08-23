mailMerge
=========

AppleScript replacement for mail merge feature missing from Pages 5

## Usage

Create a Numbers document with one sheet, containing one table, with one header row and no header columns. Fill with data.

Create a Pages document (if it has body text, almost certainly ending in a page break), and write "%Column Name%" to reference data in the column from your Numbers table whose first (header) cell is "Column Name".

Open both documents. Run script.

## Examples

Provided are Pages document "Example Practice Chart" and Numbers document "Example Students". Try opening them and running the script to see an example of what can be accomplished with mail merge.
