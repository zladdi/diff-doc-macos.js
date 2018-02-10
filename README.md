# diff-doc-macos.js

A partial port of TortoiseSVN/TortoiseGit diff-doc.js to Open Scripting Architecture (OSA) on macOS.
For the source script, see https://github.com/TortoiseGit/TortoiseGit/blob/master/contrib/diff-scripts/diff-doc.js.

The code in is distributed under the GNU General Public License. 

## Prerequisites

* Requires OS X Yosemite or later.
* Microsoft Word needs to be installed. Contrary to the source script, this port
does not support using OpenOffice (or LibreOffice) to perform the comparison. 

## Usage 

`diff-doc-macos.js <absolute-path-to-base.doc> <absolute-path-to-new.doc>`

## Known issues

* Relative paths to documents are not supported 
* Newer versions of Microsoft Office apps are sandboxed. This leads to the annoying
"Grant File Access" dialog to pop up for each of the documents to be compared in cases
where Word does not have permission to access the respective file already.

## Future work

Create a formula for [`brew`](https://github.com/Homebrew)
